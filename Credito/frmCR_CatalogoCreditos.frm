VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_CatalogoCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de Créditos"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   HelpContextID   =   3005
   Icon            =   "frmCR_CatalogoCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   11745
   Begin XtremeSuiteControls.CheckBox chkFiltrarAutoGestion 
      Height          =   255
      Left            =   7200
      TabIndex        =   93
      Top             =   480
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Filtrar Auto Gestionables"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   74
      ToolTipText     =   "Destinos o Plan de Inversión"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Destinos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":6852
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   11775
      _Version        =   1441793
      _ExtentX        =   20770
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
      ItemCount       =   4
      Item(0).Caption =   "Parámetros"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tcParametros"
      Item(1).Caption =   "Asignaciones"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGridAsg"
      Item(2).Caption =   "Rango:  Montos, Tasas, Plazos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tcRangos"
      Item(3).Caption =   "Cuentas"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "lswCuentas"
      Item(3).Control(1)=   "cmdCuentas"
      Item(3).Control(2)=   "cmdTabla"
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   5055
         Left            =   -70000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1441793
         _ExtentX        =   20770
         _ExtentY        =   8916
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcParametros 
         Height          =   6015
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   11775
         _Version        =   1441793
         _ExtentX        =   20770
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
         ItemCount       =   6
         Item(0).Caption =   "Bloque No.1"
         Item(0).ControlCount=   6
         Item(0).Control(0)=   "GroupBox1"
         Item(0).Control(1)=   "lswGarantias"
         Item(0).Control(2)=   "lswEstados"
         Item(0).Control(3)=   "scTitle(0)"
         Item(0).Control(4)=   "scTitle(1)"
         Item(0).Control(5)=   "scTitle(2)"
         Item(1).Caption =   "Bloque No.2"
         Item(1).ControlCount=   32
         Item(1).Control(0)=   "cboLiqTasa"
         Item(1).Control(1)=   "cboTipoRefun"
         Item(1).Control(2)=   "cboMoneda"
         Item(1).Control(3)=   "cboFactorCalculo"
         Item(1).Control(4)=   "cboMetodoCancelaCuotas"
         Item(1).Control(5)=   "cboRequisitos"
         Item(1).Control(6)=   "txtPorcRefun"
         Item(1).Control(7)=   "txtDias"
         Item(1).Control(8)=   "txtNumOperaciones"
         Item(1).Control(9)=   "txtPorcCancelacion"
         Item(1).Control(10)=   "txtAnticipoMesesPenalizados"
         Item(1).Control(11)=   "txtLiqValor"
         Item(1).Control(12)=   "dtpFechaCorte"
         Item(1).Control(13)=   "chkCodigoAlterno"
         Item(1).Control(14)=   "chkPermite_PersonaEnCbrJud"
         Item(1).Control(15)=   "chkFechaCorte"
         Item(1).Control(16)=   "txtMembresia"
         Item(1).Control(17)=   "Label7(8)"
         Item(1).Control(18)=   "Label7(7)"
         Item(1).Control(19)=   "Label7(6)"
         Item(1).Control(20)=   "Label6(8)"
         Item(1).Control(21)=   "Label6(7)"
         Item(1).Control(22)=   "Label6(6)"
         Item(1).Control(23)=   "Label6(5)"
         Item(1).Control(24)=   "Label7(4)"
         Item(1).Control(25)=   "Label6(4)"
         Item(1).Control(26)=   "Label7(1)"
         Item(1).Control(27)=   "Label7(2)"
         Item(1).Control(28)=   "Label6(1)"
         Item(1).Control(29)=   "Label6(2)"
         Item(1).Control(30)=   "Label7(0)"
         Item(1).Control(31)=   "Label6(3)"
         Item(2).Caption =   "Bloque No.3"
         Item(2).ControlCount=   7
         Item(2).Control(0)=   "txtNotas"
         Item(2).Control(1)=   "Label3(0)"
         Item(2).Control(2)=   "GroupBox3"
         Item(2).Control(3)=   "cboComite"
         Item(2).Control(4)=   "cboInstitucion"
         Item(2).Control(5)=   "Label6(0)"
         Item(2).Control(6)=   "Label4"
         Item(3).Caption =   "Bloque No.4"
         Item(3).ControlCount=   14
         Item(3).Control(0)=   "gbSinpe"
         Item(3).Control(1)=   "chkBonifica"
         Item(3).Control(2)=   "chkPago_Activa"
         Item(3).Control(3)=   "chkReadecua"
         Item(3).Control(4)=   "chkMntMax"
         Item(3).Control(5)=   "chkSupervision"
         Item(3).Control(6)=   "txtSupervisionMonto"
         Item(3).Control(7)=   "Label5(1)"
         Item(3).Control(8)=   "chkNotifica_Formaliza"
         Item(3).Control(9)=   "chkNotifica_Cancela"
         Item(3).Control(10)=   "chkEdadPension_Estudio"
         Item(3).Control(11)=   "chkEdadPension_Formalizacion"
         Item(3).Control(12)=   "Label5(2)"
         Item(3).Control(13)=   "txtAnticipo_Extraordinario_Porc"
         Item(4).Caption =   "Reservas/Revolutivos"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "gbReservas"
         Item(4).Control(1)=   "GroupBox2"
         Item(5).Caption =   "Auto Gestionables"
         Item(5).ControlCount=   13
         Item(5).Control(0)=   "gbTipoCrdWeb"
         Item(5).Control(1)=   "chkFP_POS"
         Item(5).Control(2)=   "chkFP_Web"
         Item(5).Control(3)=   "chkLineaVisibleEC"
         Item(5).Control(4)=   "chkWebSite"
         Item(5).Control(5)=   "vgDocAdjunto"
         Item(5).Control(6)=   "chkGirosPorLinea"
         Item(5).Control(7)=   "txtGiroMaxTransac"
         Item(5).Control(8)=   "chkGirosBancos"
         Item(5).Control(9)=   "txtGirosMntTraslado"
         Item(5).Control(10)=   "scTitulo"
         Item(5).Control(11)=   "Label2(1)"
         Item(5).Control(12)=   "Label2(0)"
         Begin XtremeSuiteControls.GroupBox gbSinpe 
            Height          =   1575
            Left            =   -69760
            TabIndex        =   132
            Top             =   4320
            Visible         =   0   'False
            Width           =   10935
            _Version        =   1441793
            _ExtentX        =   19288
            _ExtentY        =   2778
            _StockProps     =   79
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
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkSINPE 
               Height          =   255
               Left            =   480
               TabIndex        =   133
               Top             =   360
               Width           =   5895
               _Version        =   1441793
               _ExtentX        =   10398
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Habilita Transacciones SINPE ?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.ComboBox cboSinpe 
               Height          =   315
               Left            =   480
               TabIndex        =   135
               Top             =   720
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
            Begin XtremeSuiteControls.ComboBox cboTipoCredito 
               Height          =   330
               Left            =   7650
               TabIndex        =   154
               Top             =   720
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
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Credito (MEIC)"
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
               Left            =   7680
               TabIndex        =   155
               Top             =   480
               Width           =   2055
            End
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   0
               Left            =   3840
               TabIndex        =   134
               Top             =   720
               Width           =   1455
               _Version        =   1441793
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Código SINPE"
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
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   5175
            Left            =   120
            TabIndex        =   18
            Top             =   765
            Width           =   3495
            _Version        =   1441793
            _ExtentX        =   6165
            _ExtentY        =   9128
            _StockProps     =   79
            BackColor       =   16777215
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.CheckBox chkActivo 
               Height          =   252
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Línea Activa ?"
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
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkLineaInterna 
               Height          =   252
               Left            =   120
               TabIndex        =   20
               Top             =   480
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Línea de Cartera Interna?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRetencion 
               Height          =   252
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Línea de Retención"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRetencionSaldo 
               Height          =   252
               Left            =   360
               TabIndex        =   22
               Top             =   1080
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Mostrar Saldo-Retención?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkConvenio 
               Height          =   252
               Left            =   120
               TabIndex        =   23
               Top             =   1560
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Línea de Convenio"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkCodigoPoliza 
               Height          =   252
               Left            =   120
               TabIndex        =   24
               Top             =   1800
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Linea de Póliza"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRefundeOtraOperacion 
               Height          =   252
               Left            =   120
               TabIndex        =   25
               Top             =   2280
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Refunde otras Líneas"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkAceptaRefundicion 
               Height          =   252
               Left            =   120
               TabIndex        =   26
               Top             =   2520
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Acepta Refundiciones"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkPrimerCuota 
               Height          =   252
               Left            =   120
               TabIndex        =   27
               Top             =   3120
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Deduce Primer Cuota"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkPideCheque 
               Height          =   252
               Left            =   120
               TabIndex        =   28
               Top             =   3360
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Genera Cheque Automatico"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkCobertura 
               Height          =   252
               Left            =   120
               TabIndex        =   29
               Top             =   3600
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aplica Cobertura ?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkLineaMora 
               Height          =   252
               Left            =   120
               TabIndex        =   30
               Top             =   4200
               Width           =   2892
               _Version        =   1441793
               _ExtentX        =   5101
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Esta línea genera morosidad?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkAceptaMovCajas 
               Height          =   252
               Left            =   120
               TabIndex        =   31
               Top             =   4440
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Acepta Movimientos en Cajas?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkListaRefundibles 
               Height          =   252
               Left            =   120
               TabIndex        =   32
               Top             =   2760
               Width           =   2772
               _Version        =   1441793
               _ExtentX        =   4890
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Filtra las Líneas Refundibles"
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
               Appearance      =   16
            End
         End
         Begin XtremeSuiteControls.ListView lswGarantias 
            Height          =   5055
            Left            =   3720
            TabIndex        =   33
            Top             =   885
            Width           =   3855
            _Version        =   1441793
            _ExtentX        =   6800
            _ExtentY        =   8916
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
            MultiSelect     =   -1  'True
            HideSelection   =   0   'False
            View            =   3
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            BackColor       =   -2147483633
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ListView lswEstados 
            Height          =   5055
            Left            =   7680
            TabIndex        =   34
            Top             =   885
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7011
            _ExtentY        =   8916
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
            MultiSelect     =   -1  'True
            HideSelection   =   0   'False
            View            =   3
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            BackColor       =   -2147483633
            Appearance      =   16
         End
         Begin XtremeSuiteControls.ComboBox cboLiqTasa 
            Height          =   312
            Left            =   -64240
            TabIndex        =   35
            Top             =   3624
            Visible         =   0   'False
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2566
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         Begin XtremeSuiteControls.ComboBox cboTipoRefun 
            Height          =   312
            Left            =   -62680
            TabIndex        =   36
            Top             =   1824
            Visible         =   0   'False
            Width           =   1692
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
         Begin XtremeSuiteControls.ComboBox cboMoneda 
            Height          =   330
            Left            =   -64240
            TabIndex        =   37
            Top             =   4200
            Visible         =   0   'False
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
         Begin XtremeSuiteControls.ComboBox cboFactorCalculo 
            Height          =   312
            Left            =   -64240
            TabIndex        =   38
            Top             =   4584
            Visible         =   0   'False
            Width           =   3252
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
         Begin XtremeSuiteControls.ComboBox cboMetodoCancelaCuotas 
            Height          =   312
            Left            =   -64240
            TabIndex        =   39
            Top             =   4944
            Visible         =   0   'False
            Width           =   3252
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
         Begin XtremeSuiteControls.ComboBox cboRequisitos 
            Height          =   312
            Left            =   -64240
            TabIndex        =   40
            Top             =   5304
            Visible         =   0   'False
            Width           =   3252
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
         Begin XtremeSuiteControls.FlatEdit txtPorcRefun 
            Height          =   312
            Left            =   -64240
            TabIndex        =   41
            Top             =   1824
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "25"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDias 
            Height          =   312
            Left            =   -64240
            TabIndex        =   42
            Top             =   2544
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "30"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNumOperaciones 
            Height          =   312
            Left            =   -64240
            TabIndex        =   43
            Top             =   2904
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "1"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcCancelacion 
            Height          =   312
            Left            =   -64240
            TabIndex        =   44
            Top             =   3264
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "0"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAnticipoMesesPenalizados 
            Height          =   312
            Left            =   -61720
            TabIndex        =   45
            Top             =   3264
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "12"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtLiqValor 
            Height          =   312
            Left            =   -61720
            TabIndex        =   46
            Top             =   3624
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "5"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
            Height          =   312
            Left            =   -62680
            TabIndex        =   47
            Top             =   1344
            Visible         =   0   'False
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.CheckBox chkCodigoAlterno 
            Height          =   252
            Left            =   -68680
            TabIndex        =   48
            Top             =   624
            Visible         =   0   'False
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Envíar deducción por el código alterno de créditos  "
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPermite_PersonaEnCbrJud 
            Height          =   252
            Left            =   -68680
            TabIndex        =   49
            Top             =   984
            Visible         =   0   'False
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite nuevos créditos a Personas en Cobro Judicial  "
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkFechaCorte 
            Height          =   252
            Left            =   -69640
            TabIndex        =   50
            Top             =   1344
            Visible         =   0   'False
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicar diferente fecha de corte para la Formalización de esta linea  "
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
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtMembresia 
            Height          =   312
            Left            =   -64240
            TabIndex        =   51
            Top             =   2208
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "0"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   912
            Left            =   -69280
            TabIndex        =   67
            Top             =   4584
            Visible         =   0   'False
            Width           =   8892
            _Version        =   1441793
            _ExtentX        =   15684
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   1575
            Left            =   -69400
            TabIndex        =   84
            Top             =   2520
            Visible         =   0   'False
            Width           =   9495
            _Version        =   1441793
            _ExtentX        =   16743
            _ExtentY        =   2773
            _StockProps     =   79
            Caption         =   "Asigna una Oficina Fija a esta línea en la formalización..:"
            ForeColor       =   4210752
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkOficinaLinea 
               Height          =   252
               Left            =   960
               TabIndex        =   85
               Top             =   360
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Esta Línea utiliza una Oficina Fija"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.FlatEdit txtOficina 
               Height          =   312
               Left            =   960
               TabIndex        =   86
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1080
               Width           =   972
               _Version        =   1441793
               _ExtentX        =   1714
               _ExtentY        =   550
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
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
            Begin XtremeSuiteControls.FlatEdit txtOficinaDesc 
               Height          =   312
               Left            =   1920
               TabIndex        =   87
               Top             =   1080
               Width           =   7092
               _Version        =   1441793
               _ExtentX        =   12509
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Oficina Asociada a la Línea de Crédito:"
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
               Index           =   7
               Left            =   960
               TabIndex        =   88
               Top             =   840
               Width           =   3732
            End
         End
         Begin XtremeSuiteControls.ComboBox cboComite 
            Height          =   315
            Left            =   -67000
            TabIndex        =   89
            Top             =   630
            Visible         =   0   'False
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         Begin XtremeSuiteControls.ComboBox cboInstitucion 
            Height          =   315
            Left            =   -67000
            TabIndex        =   90
            Top             =   990
            Visible         =   0   'False
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         Begin XtremeSuiteControls.GroupBox gbTipoCrdWeb 
            Height          =   2055
            Left            =   -68920
            TabIndex        =   96
            Top             =   1200
            Visible         =   0   'False
            Width           =   4815
            _Version        =   1441793
            _ExtentX        =   8493
            _ExtentY        =   3625
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.RadioButton rbTipoCrdWeb 
               Height          =   252
               Index           =   0
               Left            =   360
               TabIndex        =   97
               Top             =   240
               Width           =   2052
               _Version        =   1441793
               _ExtentX        =   3619
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Crédito Directo"
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
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rbTipoCrdWeb 
               Height          =   252
               Index           =   1
               Left            =   2640
               TabIndex        =   98
               Top             =   240
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Solicitud de Crédito"
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
            End
            Begin XtremeSuiteControls.CheckBox chkRefundeAuto 
               Height          =   255
               Left            =   720
               TabIndex        =   99
               Top             =   720
               Width           =   3975
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Refunde Automáticamente Línea Actual?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRefundeAumentaBase 
               Height          =   255
               Left            =   720
               TabIndex        =   100
               Top             =   1080
               Width           =   3975
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aumenta la Base, Monto Refundido?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.FlatEdit txtGiroMinimo 
               Height          =   315
               Left            =   2880
               TabIndex        =   101
               Top             =   1680
               Width           =   1815
               _Version        =   1441793
               _ExtentX        =   3201
               _ExtentY        =   556
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "0"
               Alignment       =   1
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   102
               Top             =   1680
               Width           =   2175
               _Version        =   1441793
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Giro Mínimo Permitido"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               WordWrap        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.CheckBox chkFP_POS 
            Height          =   255
            Left            =   -63880
            TabIndex        =   103
            Top             =   600
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Activa en el POS como forma de pago?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkFP_Web 
            Height          =   255
            Left            =   -63880
            TabIndex        =   104
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Activa en la App/Web como forma de pago?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkLineaVisibleEC 
            Height          =   255
            Left            =   -68920
            TabIndex        =   105
            Top             =   600
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Es visible desde el Estado de Cuenta ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkWebSite 
            Height          =   255
            Left            =   -68920
            TabIndex        =   106
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Activa Trámite en App/Web?"
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
            Appearance      =   16
         End
         Begin FPSpreadADO.fpSpread vgDocAdjunto 
            Height          =   2295
            Left            =   -68920
            TabIndex        =   107
            Top             =   3720
            Visible         =   0   'False
            Width           =   8895
            _Version        =   524288
            _ExtentX        =   15690
            _ExtentY        =   4048
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            DisplayRowHeaders=   0   'False
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
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "frmCR_CatalogoCreditos.frx":6E6E
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.CheckBox chkGirosPorLinea 
            Height          =   495
            Left            =   -63880
            TabIndex        =   108
            Top             =   1320
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Activar Giros Maximos por Línea versus Disponible por Garantía?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtGiroMaxTransac 
            Height          =   315
            Left            =   -63640
            TabIndex        =   109
            Top             =   1920
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkGirosBancos 
            Height          =   495
            Left            =   -63880
            TabIndex        =   110
            Top             =   2280
            Visible         =   0   'False
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Activar Traslado Automático a Bancos si el monto a girar es menor a"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtGirosMntTraslado 
            Height          =   315
            Left            =   -63640
            TabIndex        =   111
            Top             =   2880
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.GroupBox gbReservas 
            Height          =   2775
            Left            =   -69520
            TabIndex        =   115
            Top             =   480
            Visible         =   0   'False
            Width           =   10455
            _Version        =   1441793
            _ExtentX        =   18441
            _ExtentY        =   4895
            _StockProps     =   79
            Caption         =   "Configuración de Reservas derivadas de la Operación...:"
            ForeColor       =   4210752
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
            Begin XtremeSuiteControls.CheckBox chkReserva_Aplica 
               Height          =   252
               Left            =   960
               TabIndex        =   116
               Top             =   360
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aplica Reserva?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkReserva_Flat 
               Height          =   252
               Left            =   960
               TabIndex        =   117
               Top             =   720
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "El Tipo de Reserva es Facial FLAT?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkReserva_Mora 
               Height          =   252
               Left            =   960
               TabIndex        =   118
               Top             =   1080
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Aplica Reserva a Morosidad?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.FlatEdit txtReserva_Plan 
               Height          =   312
               Left            =   960
               TabIndex        =   119
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   1680
               Width           =   972
               _Version        =   1441793
               _ExtentX        =   1714
               _ExtentY        =   550
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
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
            Begin XtremeSuiteControls.FlatEdit txtReserva_PlanDesc 
               Height          =   312
               Left            =   1920
               TabIndex        =   120
               Top             =   1680
               Width           =   6972
               _Version        =   1441793
               _ExtentX        =   12298
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtReserva_MontoMin 
               Height          =   312
               Left            =   960
               TabIndex        =   121
               Top             =   2400
               Width           =   2292
               _Version        =   1441793
               _ExtentX        =   4043
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Plan asociado a la Reserva:"
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
               Index           =   4
               Left            =   960
               TabIndex        =   123
               Top             =   1440
               Width           =   2172
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Monto mínimo de la Reserva:"
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
               Index           =   5
               Left            =   960
               TabIndex        =   122
               Top             =   2160
               Width           =   2532
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   2535
            Left            =   -69400
            TabIndex        =   124
            Top             =   3360
            Visible         =   0   'False
            Width           =   10215
            _Version        =   1441793
            _ExtentX        =   18018
            _ExtentY        =   4471
            _StockProps     =   79
            Caption         =   "Configuración para Línea Revolutiva...:"
            ForeColor       =   4210752
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
            Begin XtremeSuiteControls.CheckBox chkRevLinea 
               Height          =   252
               Left            =   240
               TabIndex        =   125
               Top             =   480
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Línea de Crédito Revolutiva?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRevTopeRetiros 
               Height          =   252
               Left            =   240
               TabIndex        =   126
               Top             =   840
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Utiliza el Tope de Retiros Máximos?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRevEstudio 
               Height          =   252
               Left            =   240
               TabIndex        =   127
               Top             =   1320
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Utiliza el Programa de Estudios/Convenios?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.CheckBox chkRevPlanAhorros 
               Height          =   252
               Left            =   240
               TabIndex        =   128
               Top             =   1680
               Width           =   3972
               _Version        =   1441793
               _ExtentX        =   7006
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Enlazado con Plan de Ahorros?"
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
               Appearance      =   16
            End
            Begin XtremeSuiteControls.FlatEdit txtRevPlanAhorro 
               Height          =   312
               Left            =   1320
               TabIndex        =   129
               ToolTipText     =   "Presione F4 para Consultar"
               Top             =   2160
               Width           =   972
               _Version        =   1441793
               _ExtentX        =   1714
               _ExtentY        =   550
               _StockProps     =   77
               ForeColor       =   0
               BackColor       =   -2147483643
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
            Begin XtremeSuiteControls.FlatEdit txtRevPlanDesc 
               Height          =   312
               Left            =   2280
               TabIndex        =   130
               Top             =   2160
               Width           =   6732
               _Version        =   1441793
               _ExtentX        =   11874
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Plan"
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
               Left            =   360
               TabIndex        =   131
               Top             =   2160
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.CheckBox chkBonifica 
            Height          =   255
            Left            =   -69400
            TabIndex        =   137
            Top             =   1560
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica Bonificacion al Movimiento ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkPago_Activa 
            Height          =   255
            Left            =   -69400
            TabIndex        =   138
            Top             =   1920
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Habilita Pago de la Operación"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkReadecua 
            Height          =   255
            Left            =   -69400
            TabIndex        =   139
            Top             =   2280
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Permite Readecuación de la Operación"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkMntMax 
            Height          =   255
            Left            =   -69400
            TabIndex        =   140
            Top             =   2640
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Revisa monto máximo"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkSupervision 
            Height          =   255
            Left            =   -69400
            TabIndex        =   141
            Top             =   3240
            Visible         =   0   'False
            Width           =   5055
            _Version        =   1441793
            _ExtentX        =   8916
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Control de Supervisión"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtSupervisionMonto 
            Height          =   330
            Left            =   -69400
            TabIndex        =   142
            Top             =   3600
            Visible         =   0   'False
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkNotifica_Formaliza 
            Height          =   255
            Left            =   -69400
            TabIndex        =   145
            Top             =   600
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notificar al Cliente: Al Formalizar el crédito ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkNotifica_Cancela 
            Height          =   255
            Left            =   -69400
            TabIndex        =   146
            Top             =   960
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
            _ExtentX        =   9340
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notificar al Cliente: Cuando el Crédito es cancelado?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkEdadPension_Estudio 
            Height          =   255
            Left            =   -64000
            TabIndex        =   150
            Top             =   1560
            Visible         =   0   'False
            Width           =   5535
            _Version        =   1441793
            _ExtentX        =   9763
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valida Edad de Pensión en el Estudio de Crédito ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkEdadPension_Formalizacion 
            Height          =   255
            Left            =   -64000
            TabIndex        =   151
            Top             =   1920
            Visible         =   0   'False
            Width           =   5535
            _Version        =   1441793
            _ExtentX        =   9763
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valida Edad de Pensión en la Formalización del crédito ?"
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
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtAnticipo_Extraordinario_Porc 
            Height          =   330
            Left            =   -64000
            TabIndex        =   152
            Top             =   3480
            Visible         =   0   'False
            Width           =   615
            _Version        =   1441793
            _ExtentX        =   1085
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   375
            Index           =   2
            Left            =   -63160
            TabIndex        =   153
            Top             =   3480
            Visible         =   0   'False
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "% Permitido para Anticipo Extrordinario al Principal"
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
         Begin XtremeSuiteControls.Label Label5 
            Height          =   375
            Index           =   1
            Left            =   -67360
            TabIndex        =   143
            Top             =   3600
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Monto Base para Supervisión"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   -61720
            TabIndex        =   114
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Giro Máximo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   -61720
            TabIndex        =   113
            Top             =   2880
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Traslado a Bancos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption scTitulo 
            Height          =   375
            Left            =   -68920
            TabIndex        =   112
            Top             =   3360
            Visible         =   0   'False
            Width           =   8895
            _Version        =   1441793
            _ExtentX        =   15684
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Lista de documentación requerida para los créditos y solicitudes en línea (Adjuntos)"
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
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Comité Resolutor"
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
            Left            =   -69160
            TabIndex        =   92
            Top             =   630
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Entidad Recaudadora"
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
            Left            =   -69130
            TabIndex        =   91
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
         End
         Begin XtremeShortcutBar.ShortcutCaption scTitle 
            Height          =   375
            Index           =   2
            Left            =   7680
            TabIndex        =   71
            Top             =   360
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Estados de la Persona"
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
         Begin XtremeShortcutBar.ShortcutCaption scTitle 
            Height          =   375
            Index           =   1
            Left            =   3720
            TabIndex        =   70
            Top             =   360
            Width           =   3975
            _Version        =   1441793
            _ExtentX        =   7011
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Garantías"
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
         Begin XtremeShortcutBar.ShortcutCaption scTitle 
            Height          =   372
            Index           =   0
            Left            =   0
            TabIndex        =   69
            Top             =   360
            Width           =   3732
            _Version        =   1441793
            _ExtentX        =   6583
            _ExtentY        =   656
            _StockProps     =   14
            Caption         =   "Características"
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
         Begin VB.Label Label3 
            Caption         =   "Notas"
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
            Left            =   -69244
            TabIndex        =   68
            Top             =   4320
            Visible         =   0   'False
            Width           =   2376
         End
         Begin VB.Label Label6 
            Caption         =   "Número de Operaciones Activas por Persona"
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
            Index           =   3
            Left            =   -69604
            TabIndex        =   66
            Top             =   2880
            Visible         =   0   'False
            Width           =   4212
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "% del"
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
            Index           =   0
            Left            =   -63364
            TabIndex        =   65
            Top             =   1800
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Permitir Refundiciones si a cancelado o transcurrido el"
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
            Index           =   2
            Left            =   -69604
            TabIndex        =   64
            Top             =   1800
            Visible         =   0   'False
            Width           =   4572
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Divisa"
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
            Index           =   1
            Left            =   -69604
            TabIndex        =   63
            Top             =   4200
            Visible         =   0   'False
            Width           =   2052
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "dias"
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
            Index           =   2
            Left            =   -63364
            TabIndex        =   62
            Top             =   2520
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Eliminar Solicitudes Pendientes Ingresadas despues de"
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
            Index           =   1
            Left            =   -69604
            TabIndex        =   61
            Top             =   2520
            Visible         =   0   'False
            Width           =   4812
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Aumento de Tasas en Caso de Renuncia"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   -69610
            TabIndex        =   60
            Top             =   3600
            Visible         =   0   'False
            Width           =   5295
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor"
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
            Index           =   4
            Left            =   -62404
            TabIndex        =   59
            Top             =   3600
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje de Comision x Cancelación Anticipada"
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
            Index           =   5
            Left            =   -69604
            TabIndex        =   58
            ToolTipText     =   "Genera una Comision x Pago en Cajas (Cancelando la Operación)"
            Top             =   3240
            Visible         =   0   'False
            Width           =   3852
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Factor de Cálculo de Intereses / Tipo de Amortización"
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
            Index           =   6
            Left            =   -69604
            TabIndex        =   57
            Top             =   4560
            Visible         =   0   'False
            Width           =   4452
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel de Aplicación de Requisitos (Formalización)"
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
            Index           =   7
            Left            =   -69604
            TabIndex        =   56
            Top             =   5280
            Visible         =   0   'False
            Width           =   3852
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Metodo de Cancelacion de Cuotas"
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
            Index           =   8
            Left            =   -69604
            TabIndex        =   55
            Top             =   4920
            Visible         =   0   'False
            Width           =   3852
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Meses Penalizados"
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
            Index           =   6
            Left            =   -63364
            TabIndex        =   54
            Top             =   3240
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Meses de membresía para optar por esta línea de crédito"
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
            Index           =   7
            Left            =   -69604
            TabIndex        =   53
            Top             =   2184
            Visible         =   0   'False
            Width           =   4812
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
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
            Index           =   8
            Left            =   -63364
            TabIndex        =   52
            Top             =   2184
            Visible         =   0   'False
            Width           =   852
         End
      End
      Begin XtremeSuiteControls.TabControl tcRangos 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1441793
         _ExtentX        =   20770
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
         SelectedItem    =   1
         Item(0).Caption =   "General"
         Item(0).ControlCount=   13
         Item(0).Control(0)=   "chkRngTasaDestino"
         Item(0).Control(1)=   "chkRngTBP"
         Item(0).Control(2)=   "cboMoraTipo"
         Item(0).Control(3)=   "txtTasaMora"
         Item(0).Control(4)=   "txtRngPuntosAdicionalesTBP"
         Item(0).Control(5)=   "Label7(5)"
         Item(0).Control(6)=   "Label7(3)"
         Item(0).Control(7)=   "Label9(0)"
         Item(0).Control(8)=   "Label9(1)"
         Item(0).Control(9)=   "Label9(2)"
         Item(0).Control(10)=   "chkTasaFija_TBP_Apl"
         Item(0).Control(11)=   "txtTasaFija_TBP_Pts"
         Item(0).Control(12)=   "txtTasaFija_Plazo"
         Item(1).Caption =   "Rangos Base"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "vGrid"
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   5415
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   11415
            _Version        =   524288
            _ExtentX        =   20135
            _ExtentY        =   9551
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
            SpreadDesigner  =   "frmCR_CatalogoCreditos.frx":74D0
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.CheckBox chkRngTasaDestino 
            Height          =   255
            Left            =   -69670
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8276
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Utilizar Tasa de Interes del Destino  "
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
            TextAlignment   =   1
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkRngTBP 
            Height          =   255
            Left            =   -69670
            TabIndex        =   11
            Top             =   1080
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8276
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Utilizar Método de Tasas sobre Tasa Básica Pasiva  "
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
            TextAlignment   =   1
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboMoraTipo 
            Height          =   312
            Left            =   -64996
            TabIndex        =   12
            Top             =   2280
            Visible         =   0   'False
            Width           =   3252
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
         Begin XtremeSuiteControls.FlatEdit txtTasaMora 
            Height          =   330
            Left            =   -65716
            TabIndex        =   13
            Top             =   2280
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
            _ExtentY        =   582
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
            Text            =   "5"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRngPuntosAdicionalesTBP 
            Height          =   330
            Left            =   -65716
            TabIndex        =   14
            Top             =   1560
            Visible         =   0   'False
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
            _ExtentY        =   582
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
            Text            =   "30"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkTasaFija_TBP_Apl 
            Height          =   255
            Left            =   -69670
            TabIndex        =   94
            Top             =   3000
            Visible         =   0   'False
            Width           =   4695
            _Version        =   1441793
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aplica Tasa Fija (Pts Fijos) Sobre Tasa Básica Pasiva?"
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
            TextAlignment   =   1
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtTasaFija_TBP_Pts 
            Height          =   330
            Left            =   -65710
            TabIndex        =   95
            Top             =   3510
            Visible         =   0   'False
            Width           =   735
            _Version        =   1441793
            _ExtentX        =   1291
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTasaFija_Plazo 
            Height          =   330
            Left            =   -65710
            TabIndex        =   136
            Top             =   4230
            Visible         =   0   'False
            Width           =   735
            _Version        =   1441793
            _ExtentX        =   1291
            _ExtentY        =   582
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
            Text            =   "0"
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   375
            Index           =   2
            Left            =   -64840
            TabIndex        =   149
            Top             =   4200
            Visible         =   0   'False
            Width           =   3015
            _Version        =   1441793
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "meses"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   375
            Index           =   1
            Left            =   -68920
            TabIndex        =   148
            Top             =   4200
            Visible         =   0   'False
            Width           =   3015
            _Version        =   1441793
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Plazo de Vigencias para Tasa Fija"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   375
            Index           =   0
            Left            =   -68920
            TabIndex        =   147
            Top             =   3600
            Visible         =   0   'False
            Width           =   3015
            _Version        =   1441793
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Puntos Fijos Sobre Tasa Básica Pasiva"
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
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Puntos Adicionales Sobre Tasa Básica Pasiva"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   -69400
            TabIndex        =   16
            Top             =   1530
            Visible         =   0   'False
            Width           =   3495
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Tasa por Morosidad (Pts Adicionales ó Porcentaje s/tasa vigente)"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Index           =   5
            Left            =   -69880
            TabIndex        =   15
            Top             =   2136
            Visible         =   0   'False
            Width           =   3852
         End
      End
      Begin XtremeSuiteControls.PushButton cmdCuentas 
         Height          =   492
         Left            =   -60640
         TabIndex        =   9
         Top             =   5520
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cuentas"
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
         Picture         =   "frmCR_CatalogoCreditos.frx":8597
         ImageAlignment  =   0
      End
      Begin FPSpreadADO.fpSpread vGridAsg 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   11655
         _Version        =   524288
         _ExtentX        =   20558
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
         MaxCols         =   497
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_CatalogoCreditos.frx":8CB0
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdTabla 
         Height          =   492
         Left            =   -70000
         TabIndex        =   72
         Top             =   5520
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Tabla"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCR_CatalogoCreditos.frx":9D4C
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10440
      Top             =   840
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9720
      TabIndex        =   1
      Top             =   840
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigoCorriente 
      Height          =   312
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Height          =   312
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11451
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   2160
      TabIndex        =   73
      Top             =   480
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Reportes del catalogo de préstamos"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OrdCodigo"
                  Text            =   "Ordenado por Código"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OrdDes"
                  Text            =   "Ordenado por Descripción"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Separador1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DetalladoCod"
                  Text            =   "Detallado - Ordenado por Código "
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DetalladoDesc"
                  Text            =   "Detallado - Ordenado por Descripción"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "separador2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "linea"
                  Text            =   "Detalle de este línea"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "separador3"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repCuentasCod"
                  Text            =   "Cuentas Contables - Por Código"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Consulta el catalogo de préstamos"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Sale de esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   1
      Left            =   1440
      TabIndex        =   75
      ToolTipText     =   "Requisitos"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Requisitos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":A52B
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   2
      Left            =   2640
      TabIndex        =   76
      ToolTipText     =   "Cargos Adicionales"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Cargos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":AC52
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   3
      Left            =   3840
      TabIndex        =   77
      ToolTipText     =   "Recursos Presupuestarios"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Recursos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":B339
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   4
      Left            =   5280
      TabIndex        =   78
      ToolTipText     =   "Carteras de Crédito y Cobro"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Cartera"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":BA59
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   5
      Left            =   6480
      TabIndex        =   79
      ToolTipText     =   "Tipos de Garantías"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Garantías"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":C172
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   6
      Left            =   7680
      TabIndex        =   81
      ToolTipText     =   "Niveles de Resolución"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Autorizados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":C879
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   7
      Left            =   9120
      TabIndex        =   82
      ToolTipText     =   "Prioridades de Deducción por Línea"
      Top             =   36
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Prioridad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":CFA0
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccesoDirecto 
      Height          =   312
      Index           =   8
      Left            =   10320
      TabIndex        =   83
      ToolTipText     =   "Prioridades de Deducción por Línea"
      Top             =   40
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Copiar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmCR_CatalogoCreditos.frx":D5C4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkFiltraActivas 
      Height          =   255
      Left            =   9480
      TabIndex        =   144
      Top             =   480
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Filtrar Activas"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   382
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Width           =   12732
      _Version        =   1441793
      _ExtentX        =   22458
      _ExtentY        =   674
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
      VisualTheme     =   6
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   252
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   852
   End
End
Attribute VB_Name = "frmCR_CatalogoCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strConsulta As String, vPaso As Boolean
Dim intEdita As Integer 'Indica si está modificando
Dim txtBox As TextBox, vScroll As Boolean, mCodigoAlterno As String



Private Function fxValidaDatos() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim vMensaje As String

fxValidaDatos = True
vMensaje = ""

Select Case ""
    Case Trim(txtCodigoCorriente)
      vMensaje = vMensaje & vbCrLf & " - Código de Línea no esválido..."
    Case Trim(txtDescripcion)
      vMensaje = vMensaje & vbCrLf & " - Descripción no válida..."
    Case Trim(cboComite.Text)
      vMensaje = vMensaje & vbCrLf & " - Comité de Evaluación no es válido..."
    Case Trim(cboInstitucion.Text)
      vMensaje = vMensaje & vbCrLf & " - Entidad Deductora no es válida..."
End Select

If Len(Trim(txtCodigoCorriente)) > 4 Then vMensaje = vMensaje & vbCrLf & " - Código Corriente [excede letras]Inválido..."
'If Len(Trim(mCodigoAlterno)) > 4 Then vMensaje = vMensaje & vbCrLf & " - Código de Atraso [excede letras]Inválido..."


If Not IsNumeric(txtAnticipo_Extraordinario_Porc) Then
      vMensaje = vMensaje & vbCrLf & " - El % de Anticipo Extraordinario, no es válido ..."
Else
   If CCur(txtAnticipo_Extraordinario_Porc) < 0 Or CCur(txtAnticipo_Extraordinario_Porc) > 100 Then
      vMensaje = vMensaje & vbCrLf & " - El % de Anticipo Extraordinario, no es válido ..."
   End If
End If


If Not IsNumeric(txtPorcCancelacion) Then
      vMensaje = vMensaje & vbCrLf & " - El % de Comisión x Cancelación (Anticipo), no es válido ..."
Else
   If CCur(txtPorcCancelacion) < 0 Or CCur(txtPorcCancelacion) > 100 Then
      vMensaje = vMensaje & vbCrLf & " - El % de Comisión x Cancelación (Anticipo), no es válido ..."
   End If
End If

If Not IsNumeric(txtPorcRefun) Then
      vMensaje = vMensaje & vbCrLf & " - El % de Amortización para Permitir Refundiciones, no es válido ..."
Else
   If CCur(txtPorcRefun) < 0 Or CCur(txtPorcRefun) > 100 Then
      vMensaje = vMensaje & vbCrLf & " - El % de Amortización para Permitir Refundiciones, no es válido ..."
   End If
End If

If Not IsNumeric(txtTasaMora) Then
      vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales para Morosidad, no es válido ..."
Else
   If CCur(txtTasaMora) < 0 Or CCur(txtTasaMora) > 100 Then
      vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales para Morosidad, no es válido ..."
   End If
End If

If Not IsNumeric(txtRngPuntosAdicionalesTBP) Then
      vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales Sobre TBP, no es válido ..."
Else
   If CCur(txtRngPuntosAdicionalesTBP) < 0 Or CCur(txtRngPuntosAdicionalesTBP) > 100 Then
      vMensaje = vMensaje & vbCrLf & " - Los Puntos Adicionales Sobre TBP, no es válido ..."
   End If
End If


If Not IsNumeric(txtReserva_MontoMin.Text) Then
      vMensaje = vMensaje & vbCrLf & " - El monto de reserva mínima, no es válido ..."
End If


If Not IsNumeric(txtGiroMinimo.Text) Then
      vMensaje = vMensaje & vbCrLf & " - El monto para Giros Mínimos, no es válido ..."
End If

If Not IsNumeric(txtGiroMaxTransac.Text) Then
      vMensaje = vMensaje & vbCrLf & " - El monto para Giros Maximos por Transacción, , no es válido ..."
End If

If Not IsNumeric(txtGirosMntTraslado.Text) Then
      vMensaje = vMensaje & vbCrLf & " - El monto para Giros a Trasladar a Bancos Automatico, no es válido ..."
End If



If chkOficinaLinea.Value = vbChecked And Trim(txtOficina.Text) = "" Then
      vMensaje = vMensaje & vbCrLf & " - La Oficina Fija no fue especificada para esta línea..."
End If

If intEdita = 0 Then
 'Insertar
 strSQL = "select isnull(count(*),0) as Existe from catalogo where codigo = '" & Trim(txtCodigoCorriente) & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe > 0 Then vMensaje = vMensaje & vbCrLf & " - El código corriente ya existe..."
 rs.Close
End If


If Len(vMensaje) > 1 Then
 fxValidaDatos = False
 MsgBox vMensaje, vbCritical
End If

End Function

Private Sub btnAccesoDirecto_Click(Index As Integer)

Dim pForm As String, pModal As Integer

pModal = 0

Select Case Index
  Case 0 'Destinos
    pForm = "frmCR_CatalogoDestinos"
  Case 1 'Requisitos
    pForm = "frmCR_CatalogoRequisitos"
  Case 2 'Cargos
    pForm = "frmCR_CatalogoCargos"
  Case 3 'Recursos
    pForm = "frmCR_CatalogoGrupos"
  Case 4 'Cartera de Credito
    pForm = "frmCO_Cartera"
  Case 5 'Garantias
    pForm = "frmCR_CatalogoGarantias"
  Case 6 'Autorizados, Niveles de Resolución
    pForm = "frmCR_Niveles"
  Case 7 'Prioridad de Dedución por Linea
    pForm = "frmCR_Prioridad"
  Case 8 'Copiar
    pForm = "frmCR_CatalogoCopia"

    pModal = 1
End Select

Call sbFormsCall(pForm, , , , False, Me)

End Sub


Private Sub cboComite_Change()
'   chkRefundeOtraOperacion.SetFocus
End Sub

Private Sub chkRetencion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica que no existan operaciones activas
' ya que esto representa un cambio procedimental
' que puede traer diferencias en los auxiliares por su forma de comportamiento

'Si no esta activado el paso, sale de procedimiento porque se encuentra en proceso de carga
If vPaso Then Exit Sub

'Las Retenciones y Polizas son compatibles procedimentalmente
If chkCodigoPoliza.Value = vbChecked Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select isnull(count(*),0) as Casos from reg_Creditos where estado = 'A' and saldo > 0" _
       & " and codigo = '" & txtCodigoCorriente & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 1 Then
  Me.MousePointer = vbDefault
  MsgBox "Existen Operaciones Activas con el Estado Contrario al Marcado, por este motivo no se permite es cambio", vbExclamation
  
  vPaso = True
      chkRetencion.Value = IIf((chkRetencion.Value = vbChecked), vbUnchecked, vbChecked)
  vPaso = False
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkCodigoPoliza_Click()
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica que no existan operaciones activas
' ya que esto representa un cambio procedimental
' que puede traer diferencias en los auxiliares por su forma de comportamiento

'Si no esta activado el paso, sale de procedimiento porque se encuentra en proceso de carga
If vPaso Then Exit Sub

'Las Retenciones y Polizas son compatibles procedimentalmente
If chkRetencion.Value = vbChecked Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select isnull(count(*),0) as Casos from reg_Creditos where estado = 'A' and saldo > 0" _
       & " and codigo = '" & txtCodigoCorriente & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Casos > 1 Then
  Me.MousePointer = vbDefault
  MsgBox "Existen Operaciones Activas con el Estado Contrario al Marcado, por este motivo no se permite es cambio", vbExclamation
  vPaso = True
      chkCodigoPoliza.Value = IIf((chkCodigoPoliza.Value = vbChecked), vbUnchecked, vbChecked)
  vPaso = False
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFechaCorte_Click()
If chkFechaCorte.Value = vbChecked Then
  dtpFechaCorte.Enabled = True
Else
  dtpFechaCorte.Enabled = False
End If
End Sub


Private Sub sbCargaCuentas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

    
cmdTabla.Visible = False
    
strSQL = "Select * from vCrd_Catalogo_Cuentas where codigo = '" & txtCodigoCorriente & "'"
Call OpenRecordSet(rs, strSQL)

'Llena Cuentas
With lswCuentas
   .ListItems.Clear

If Not rs.EOF And Not rs.BOF Then
       Set itmX = .ListItems.Add(, , "NORMAL")
           itmX.ForeColor = vbBlue
           
       Set itmX = .ListItems.Add(, , "   Principal")
           itmX.SubItems(1) = rs!ctaNamort_Mask & ""
           itmX.SubItems(2) = rs!ctaNamort_Desc & ""
       
       Set itmX = .ListItems.Add(, , "   Int.Corriente")
           itmX.SubItems(1) = rs!ctaNintC_Mask & ""
           itmX.SubItems(2) = rs!ctaNintC_Desc & ""
       
       Set itmX = .ListItems.Add(, , "   Int.Moratorio")
           itmX.SubItems(1) = rs!ctaNintM_Mask & ""
           itmX.SubItems(2) = rs!ctaNintM_Desc & ""
           
       Set itmX = .ListItems.Add(, , "OPEX")
           itmX.ForeColor = vbBlue
       
       Set itmX = .ListItems.Add(, , "   Principal")
           itmX.SubItems(1) = rs!CtaOamort_Mask & ""
           itmX.SubItems(2) = rs!CtaOamort_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Int.Corriente")
           itmX.SubItems(1) = rs!ctaOintC_Mask & ""
           itmX.SubItems(2) = rs!ctaOintC_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Int.Moratorio")
           itmX.SubItems(1) = rs!ctaOintM_Mask & ""
           itmX.SubItems(2) = rs!ctaOintM_Desc & ""
           
       Set itmX = .ListItems.Add(, , "CBR.JUD.")
           itmX.ForeColor = vbBlue
           
       Set itmX = .ListItems.Add(, , "   Principal")
           itmX.SubItems(1) = rs!ctacamort_Mask & ""
           itmX.SubItems(2) = rs!CtaCamort_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Int.Corriente")
           itmX.SubItems(1) = rs!ctacintc_Mask & ""
           itmX.SubItems(2) = rs!ctaCintC_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Int.Moratorio")
           itmX.SubItems(1) = rs!ctacintm_Mask & ""
           itmX.SubItems(2) = rs!ctaCintM_Desc & ""
           
       Set itmX = .ListItems.Add(, , "COMPLEMENTARIAS")
           itmX.ForeColor = vbBlue
           
       'Cargos Anticipado e IVA
       Set itmX = .ListItems.Add(, , "   Cancelación Anticipada")
           itmX.SubItems(1) = rs!CTA_CARGOS_ANTICIPO_Mask & ""
           itmX.SubItems(2) = rs!CTA_CARGOS_ANTICIPO_Desc & ""
       
       Set itmX = .ListItems.Add(, , "   Imp.Valor Agregado")
           itmX.SubItems(1) = rs!CTA_IVA_Mask & ""
           itmX.SubItems(2) = rs!CTA_IVA_Desc & ""
       
           
       'Producto Acumulado
       Set itmX = .ListItems.Add(, , "   Prod.Acum.Cartera")
           itmX.SubItems(1) = rs!CTA_CAR_PRODUCTO_Mask & ""
           itmX.SubItems(2) = rs!CTA_CAR_PRODUCTO_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Prod.Acum.Efectos (+/-)")
           itmX.SubItems(1) = rs!CTA_PROD_ACUM_Mask & ""
           itmX.SubItems(2) = rs!CTA_PROD_ACUM_Desc & ""
           
       'Interes Cobrado por Adelantado
       Set itmX = .ListItems.Add(, , "   Int.Cbr. x Adelantado")
           itmX.SubItems(1) = rs!CTA_INT_ADELANTADO_Mask & ""
           itmX.SubItems(2) = rs!CTA_INT_ADELANTADO_Desc & ""
           
       'Producto Acumulado en Suspenso
       Set itmX = .ListItems.Add(, , "Registra Produto en Suspenso")
           itmX.ForeColor = vbBlue
          If rs!PS_REGISTRA = 1 Then
           itmX.SubItems(1) = "Sí"
          Else
           itmX.SubItems(1) = "No"
          End If
          
       Set itmX = .ListItems.Add(, , "   Prod.Susp.Deudora")
           itmX.SubItems(1) = rs!CTA_PS_DEUDORA_Mask & ""
           itmX.SubItems(2) = rs!CTA_PS_DEUDORA_Desc & ""
           
       Set itmX = .ListItems.Add(, , "   Prod.Susp.Acreedora")
           itmX.SubItems(1) = rs!CTA_PS_ACREADORA_Mask & ""
           itmX.SubItems(2) = rs!CTA_PS_ACREADORA_Desc & ""
           
       Set itmX = .ListItems.Add(, , "PUENTE")
           itmX.ForeColor = vbBlue
       
       Set itmX = .ListItems.Add(, , "   Cierre Formalización")
           itmX.SubItems(1) = rs!ctapuente_Mask & ""
           itmX.SubItems(2) = rs!ctapuente_Desc & ""
End If

End With

rs.Close

End Sub



Private Sub chkRngTasaDestino_Click()
If chkRngTasaDestino.Value = vbChecked Then
  chkRngTBP.Enabled = False
  txtRngPuntosAdicionalesTBP.Enabled = False
Else
  chkRngTBP.Enabled = True
  txtRngPuntosAdicionalesTBP.Enabled = True
End If
End Sub

Private Sub cmdCuentas_Click()

GLOBALES.gTag = txtCodigoCorriente.Text
GLOBALES.gTag2 = txtDescripcion.Text

Call sbFormsCall("frmCR_CtaCatalogo", vbModal, , , False, Me, True)
Call sbCargaCuentas

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 codigo from catalogo"
    
    If chkFiltrarAutoGestion.Value = xtpChecked Then
    
            If FlatScrollBar.Value = 1 Then
               strSQL = strSQL & " where codigo > '" & txtCodigoCorriente.Text & "' and WebSite = 1 order by codigo asc"
            Else
               strSQL = strSQL & " where codigo < '" & txtCodigoCorriente.Text & "' and WebSite = 1 order by codigo desc"
            End If
    Else
    
            If FlatScrollBar.Value = 1 Then
               strSQL = strSQL & " where codigo > '" & txtCodigoCorriente.Text & "' order by codigo asc"
            Else
               strSQL = strSQL & " where codigo < '" & txtCodigoCorriente.Text & "' order by codigo desc"
            End If
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigoCorriente.Text = rs!Codigo
      
      
        intEdita = 2
        Call sbToolBar(tlbPrincipal, "activo")
      
      txtCodigoCorriente_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 3
End Sub

Private Sub Form_Load()
Dim i As Integer

On Error GoTo vError
 
With lswGarantias.ColumnHeaders
    .Clear
    .Add , , "", 3000
End With
 
With lswEstados.ColumnHeaders
    .Clear
    .Add , , "", 3000
End With
 
 
 Call sbToolBarIconos(tlbPrincipal, False)
 vModulo = 3
 
 vGrid.AppearanceStyle = fxGridStyle
 vGridAsg.AppearanceStyle = vGrid.AppearanceStyle
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 Call Formularios(Me)    'Carga los tags de los controles, utiliza la variable global de arriba
                                 
 '0 = Inserta ,1 = Edita, 2 = Consulta
 intEdita = 2
 
 With lswCuentas.ColumnHeaders
    .Add , , "Rubro", 3500
    .Add , , "Cuenta", 2800, vbCenter
    .Add , , "Descripción", 4200
 End With
 
 cboSinpe.Clear
 cboSinpe.AddItem "Trámite Interbancario"
 cboSinpe.ItemData(cboSinpe.ListCount - 1) = CStr(1)
 cboSinpe.AddItem "No agregar motivo"
 cboSinpe.ItemData(cboSinpe.ListCount - 1) = CStr(3)
 
 cboSinpe.Text = "No agregar motivo"


 cboTipoCredito.Clear
 cboTipoCredito.AddItem "Credito"
 cboTipoCredito.ItemData(cboTipoCredito.ListCount - 1) = "C"
 cboTipoCredito.AddItem "Microcredito"
 cboTipoCredito.ItemData(cboTipoCredito.ListCount - 1) = "M"
 
 cboTipoCredito.Text = "Credito"
 
cboRequisitos.AddItem "Línea"
cboRequisitos.AddItem "Garantía"
cboRequisitos.Text = "Garantía"

cboMoraTipo.AddItem "Puntos Adicionales"
cboMoraTipo.AddItem "Porcentaje s/Tasa Vigente"
cboMoraTipo.AddItem "No Calcula Int.Moratorio"
cboMoraTipo.AddItem "Tasa Fija"
cboMoraTipo.Text = "Porcentaje s/Tasa Vigente"

Call sbCrd_Factor_Calculo(cboFactorCalculo)

cboMetodoCancelaCuotas.Clear
cboMetodoCancelaCuotas.AddItem "Horizontal"
cboMetodoCancelaCuotas.AddItem "Vertical"
cboMetodoCancelaCuotas.Text = "Horizontal"

cboTipoRefun.AddItem "01 - Plazo"
cboTipoRefun.AddItem "02 - Monto"
cboTipoRefun.Text = "01 - Plazo"
 
cboLiqTasa.AddItem "Aumento de Puntos"
cboLiqTasa.AddItem "Tasa Fija"
cboLiqTasa.Text = "Aumento de Puntos"
  
 Call sbToolBar(tlbPrincipal, "nuevo")
 
'Inicializa Formularios
tcMain.Item(0).Selected = True
tcParametros.Item(0).Selected = True

For i = 0 To tcMain.ItemCount - 1
  tcMain.Item(i).Enabled = False
Next i

cmdTabla.Enabled = False
cmdCuentas.Enabled = False
  
 Call RefrescaTags(Me) 'Apaga los botones a los que el usuario no tiene derechos
 
 vGrid.Enabled = cmdTabla.Enabled
 
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbAsignaciones(Optional pSheet As Integer = 1)
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

With vGridAsg

vPaso = True

.Sheet = pSheet

Select Case pSheet
  Case 1 'Carga Destinos
        strSQL = "select R.*,A.codigo as Existe" _
               & " from Catalogo_Destinos R left Join catalogo_destinosAsg A " _
               & " on R.cod_destino = A.cod_destino and A.codigo = '" _
               & txtCodigoCorriente.Text & "' order by existe desc,R.cod_destino"
        Call OpenRecordSet(rs, strSQL, 0)
        
        .MaxRows = 0
        .MaxCols = 3
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!cod_destino
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
  
  Case 2 'Carga Cargos
        strSQL = "select R.*,A.codigo as Existe" _
               & " from Cargos_Adicionales R left Join Cargos_asignacion A " _
               & " on R.cod_cargo = A.cod_cargo and A.codigo = '" _
               & txtCodigoCorriente & "' order by existe desc,R.cod_cargo"
        Call OpenRecordSet(rs, strSQL, 0)
        
        .MaxRows = 0
        .MaxCols = 5
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!COD_CARGO
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Text = IIf((rs!Tipo = "P"), "Porcentual", "Monto")
          .col = 4
          .Text = CStr(Format(rs!Valor, "##,###,##0.0000"))
          .col = 5
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
  
  Case 3 'Carga Requisitos
        strSQL = "select R.*,isnull(A.opcional,0) as 'OpcionalX',A.codigo as Existe" _
               & " from Requisitos_Adicionales R left Join Requisitos_asignacion A " _
               & " on R.cod_requisito = A.cod_requisito and A.codigo = '" _
               & txtCodigoCorriente.Text & "' order by existe desc,R.cod_requisito"
        
        .MaxRows = 0
        .MaxCols = 4
        
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!COD_REQUISITO
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Value = rs!OpcionalX
          .col = 4
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
  
  Case 4 'Carga Recursos
        strSQL = "select G.*,A.codigo as Existe" _
               & " from catalogo_grupos G left Join catalogo_asignaGrp A " _
               & " on G.cod_grupo = A.cod_grupo and A.codigo = '" _
               & txtCodigoCorriente.Text & "' order by existe desc,G.cod_grupo"
        
        Call OpenRecordSet(rs, strSQL, 0)
        
        .MaxRows = 0
        .MaxCols = 3
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!Cod_Grupo
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
        
   Case 5 'Cartera de Mora
        strSQL = "select R.*,A.codigo as Existe" _
               & " from CBR_CLASIFICACION_CARTERA R left Join CBR_CLASIFICACION_DETALLE A " _
               & " on R.COD_CLASIFICACION = A.COD_CLASIFICACION and A.codigo = '" _
               & txtCodigoCorriente.Text & "' order by existe desc,R.COD_CLASIFICACION"
        
        Call OpenRecordSet(rs, strSQL, 0)
        
        .MaxRows = 0
        .MaxCols = 3
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!COD_CLASIFICACION
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
        
        
        
   Case 6 'Lista de Refundibles
        strSQL = "select R.*,A.codigo as Existe" _
               & " from CATALOGO R left Join CRD_CATALOGO_REFUNDIBLES A " _
               & " on R.CODIGO = A.COD_REFUNDIBLE and A.CODIGO= '" & txtCodigoCorriente.Text _
               & "' order by existe desc,R.codigo"
        
        Call OpenRecordSet(rs, strSQL, 0)
        
        .MaxRows = 0
        .MaxCols = 3
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .col = 1
          .Text = rs!Codigo
          .col = 2
          .Text = rs!Descripcion
          .col = 3
          .Value = IIf(IsNull(rs!Existe), 0, 1)
          rs.MoveNext
        Loop
        rs.Close
        
 End Select 'pSheet

vPaso = False

End With

Me.MousePointer = vbDefault

End Sub


Private Sub sbCargaGarantias()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select G.*,A.codigo as Existe" _
       & " from CRD_Garantia_Tipos G left Join Crd_Catalogo_Garantias A " _
       & " on G.Garantia = A.Garantia and A.codigo = '" _
       & txtCodigoCorriente & "' order by existe desc,G.Garantia"
Call OpenRecordSet(rs, strSQL, 0)

lswGarantias.ListItems.Clear
lswGarantias.Visible = True

Do While Not rs.EOF
  Set itmX = lswGarantias.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!Garantia
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      If itmX.Checked Then itmX.ForeColor = vbBlue
  rs.MoveNext
Loop
rs.Close


vPaso = False
Me.MousePointer = vbDefault

End Sub


Private Sub sbCargaEstados()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select E.*,A.codigo as Existe" _
       & " from AFI_Estados_Persona E left Join Crd_Catalogo_Estados A " _
       & " on E.cod_estado = A.cod_estado and A.codigo = '" _
       & txtCodigoCorriente & "' order by existe desc,E.cod_Estado"
Call OpenRecordSet(rs, strSQL, 0)

lswEstados.ListItems.Clear
lswEstados.Visible = True

Do While Not rs.EOF
  Set itmX = lswEstados.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!cod_estado
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close


vPaso = False
Me.MousePointer = vbDefault

End Sub


Private Sub sbCargaRangos(Optional pSheet As Integer = 1)
Dim strSQL As String

tcRangos.Item(1).Selected = True

 vGrid.Sheet = pSheet
 Select Case pSheet
  Case 1 'Rangos de Montos
        strSQL = "select consec,de,hasta,plazo,intc_soc,intm_soc,intc_nsoc,intm_nsoc" _
               & " from Rangos where codigo = '" & txtCodigoCorriente & "'"
        Call sbCargaGridFps7(vGrid, 8, strSQL, True, 1)
     
   Case 2 'Plazos
        strSQL = "select consec,desde,hasta,tasa" _
               & " from Rangos_plazo where codigo = '" & txtCodigoCorriente & "'"
        Call sbCargaGridFps7(vGrid, 4, strSQL, True, 2)
   
   Case 3 'Garantias
        strSQL = "select G.Garantia,G.Descripcion,A.*" _
               & " from crd_garantia_Tipos G inner join crd_catalogo_garantias A on G.garantia = A.garantia" _
               & " where A.codigo = '" & txtCodigoCorriente.Text & "'"
        Call sbCargaGridGarantias(10, strSQL)

 End Select

End Sub

Public Sub sbCargaGridGarantias(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer


vGrid.Sheet = 3
vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
  
    vGrid.col = i
    Select Case i
       Case 1 'Garantia
            vGrid.CellTag = rs!Garantia
            vGrid.Text = rs!Descripcion
            
            vGrid.TextTip = TextTipFixed
            vGrid.TextTipDelay = 1000
            vGrid.CellNote = "Registro : " & rs!Registro_Usuario & "[" & rs!Registro_Fecha & "]" & vbCrLf _
                           & "Actualizado: " & rs!Actualiza_usuario & "[" & rs!Actualiza_fecha & "]"
            
       Case 2 'Utiliza Tasa Garantia
            vGrid.Value = rs!utiliza_tasa_Garantia
       Case 3 'Tasa Garantia
            vGrid.Text = CStr(rs!Tasa_Garantia)
       Case 4 'Utiliza Tasa Piso
            vGrid.Value = rs!Utiliza_Tasa_Piso
       Case 5 'Tasa Piso
            vGrid.Text = CStr(rs!Tasa_Piso)
       Case 6 'Utiliza Tasa Techo
            vGrid.Value = rs!Utiliza_Tasa_Techo
       Case 7 'Tasa Techo
            vGrid.Text = CStr(rs!Tasa_Techo)
       Case 8 'Utiliza Montos Maximos
            vGrid.Value = rs!Utiliza_Maximos
       Case 9 'Montos Maximos
            vGrid.Text = Format(rs!Max_Monto, "Standard")
       Case 10 'Liquidez Minima
            vGrid.Text = CStr(rs!Liquidez_Minima)
    End Select
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub








Private Sub lswEstados_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vMovimiento As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  vMovimiento = "Registra"
  strSQL = "insert Crd_Catalogo_Estados(codigo,cod_estado,registro_usuario,registro_fecha) values('" & txtCodigoCorriente.Text _
         & "','" & Item.Tag & "','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
  vMovimiento = "Borrar"
  strSQL = "delete Crd_Catalogo_Estados where codigo = '" _
         & txtCodigoCorriente & "' and cod_estado = '" & Item.Tag & "'"

End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Estado Persona : " & Item.Text & " a la Línea :" & txtCodigoCorriente)
Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub lswGarantias_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, vMovimiento As String

If vPaso Then Exit Sub


On Error GoTo vError

If Item.Checked Then
  vMovimiento = "Registra"
  strSQL = "insert Crd_Catalogo_Garantias(codigo,Garantia,utiliza_tasa_garantia, tasa_garantia, liquidez_minima" _
         & ", utiliza_tasa_piso,utiliza_tasa_techo,utiliza_maximos" _
         & ",Tasa_piso,Tasa_Techo,Max_Monto,registro_usuario,registro_fecha,estado) values('" & txtCodigoCorriente.Text _
         & "','" & Item.Tag & "',0,0,0,0,0,0,0,0,0,'" & glogon.Usuario & "',dbo.MyGetdate(),1)"
Else
  vMovimiento = "Borrar"
  strSQL = "delete Crd_Catalogo_Garantias where codigo = '" _
         & txtCodigoCorriente & "' and Garantia = '" & Item.Tag & "'"

End If
Call ConectionExecute(strSQL)

Call Bitacora(vMovimiento, "Garantía : " & Item.Text & " a la Línea :" & txtCodigoCorriente)
Exit Sub


vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Parametros
        'Nada
    Case 1 'Asignaciones
        Call sbAsignaciones(vGridAsg.ActiveSheet)
    
    Case 2 'Rangos
        tcRangos.Item(0).Selected = True
        Call chkRngTasaDestino_Click

    Case 3 'Cuentas
        Call sbCargaCuentas
End Select

End Sub

Private Sub sbAdjuntos_Load()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

With vgDocAdjunto
    .MaxRows = 0
    .MaxCols = 4
    
    vPaso = True


    strSQL = "select R.*,isnull(A.opcional,0) as 'OpcionalX',A.COD_ADJUNTO as Existe" _
           & " from CRD_ADJUNTOS_TIPOS R left Join CRD_CATALOGO_ADJUNTOS A " _
           & " on R.COD_ADJUNTO = A.COD_ADJUNTO and A.CODIGO = '" & txtCodigoCorriente.Text _
           & "' order by existe desc,R.COD_ADJUNTO"
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .col = 1
      .Text = rs!COD_ADJUNTO
      .col = 2
      .Text = rs!Descripcion
      .col = 3
      .Value = rs!OpcionalX
      .col = 4
      .Value = IIf(IsNull(rs!Existe), 0, 1)
      rs.MoveNext
    Loop
    rs.Close

vPaso = False

End With

Me.MousePointer = vbDefault

End Sub




Private Sub tcParametros_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

'Adjuntos
If Item.Index = 4 Then
    Call sbAdjuntos_Load
End If

End Sub

Private Sub tcRangos_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Index = 1 Then
    Call sbCargaRangos(1)
End If
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

 strSQL = "select id_comite as Idx,rtrim(descripcion) as ItmX from comites"
 Call sbCbo_Llena_New(cboComite, strSQL, False, True)
 
 If cboComite.ListCount = 0 Then
    MsgBox "No existen Comités creados...(Debe Crearlos)", vbCritical
 End If
  
  
 strSQL = "exec spSys_Divisas"
 Call sbCbo_Llena_New(cboMoneda, strSQL, False, True)
  
  
 'Instituciones de Deduccion
 strSQL = "select cod_institucion as Idx,rtrim(descripcion) as Itmx from instituciones"
 Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)
 
 If cboInstitucion.ListCount = 0 Then
    MsgBox "No existen Instituciones creadas...(Debe Crearlos)", vbCritical
 End If
 
 lswGarantias.ShowBorder = True
 lswEstados.ShowBorder = True
 
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Button.Key
    Case "insertar"
       intEdita = 0
       Call sbLimpiaPantalla
       
       tcMain.Item(0).Enabled = True
       
       Call sbToolBar(tlbPrincipal, "edicion")
       
       cmdTabla.Enabled = False
       cmdCuentas.Enabled = False
        
       lswGarantias.Visible = False
       lswEstados.Visible = False
        
    Case "modificar"
         
         If Trim(txtCodigoCorriente) = "" Then
           MsgBox "Consulte el código que desea modificar y luego selecciones esta opción", vbOKOnly
           Exit Sub
         Else
            intEdita = 1
            Call sbToolBar(tlbPrincipal, "edicion")
            
            cmdTabla.Enabled = True
            cmdCuentas.Enabled = True
         End If
    
    Case "borrar"
      
      If txtCodigoCorriente <> "" Then
        If MsgBox("Está seguro que desea borrar este código", vbYesNo) = vbYes Then
          glogon.Conection.Execute "Delete catalogo where codigo = '" _
                                 & txtCodigoCorriente.Text & "'"
          Call Bitacora("Borra", "Codigo = " & Trim(txtCodigoCorriente))
        End If
      End If
    
    Case "deshacer"
        intEdita = 2
        sbLimpiaPantalla
        Call sbToolBar(tlbPrincipal, "activo")
        Call RefrescaTags(Me)
    
    Case "guardar"
        If fxValidaDatos Then     'existen todos los datos de la pantalla
         Call sbGuardaLinea
         Call sbToolBar(tlbPrincipal, "activo")
         cmdTabla.Enabled = True
         cmdCuentas.Enabled = True
         
         Call RefrescaTags(Me)
         Else
         
         If (MsgBox("Faltan datos, desea limpiar la información", vbYesNo)) = vbYes Then
            Call sbToolBar(tlbPrincipal, "activo")
            sbLimpiaPantalla
            RefrescaTags (Me)
         End If
       End If
    
    Case "consultar"
        
        sbLimpiaPantalla
        intEdita = 2
        gBusquedas.Consulta = "select codigo,descripcion,codigoa from catalogo"
        gBusquedas.Columna = strConsulta
        gBusquedas.Resultado = ""
        
        Select Case strConsulta
          Case "codigo"
                gBusquedas.Orden = "codigo"
          Case "codigoa"
                gBusquedas.Orden = "codigoa"
          Case "descripcion"
                gBusquedas.Orden = "descripcion"
          Case Else
                gBusquedas.Orden = "codigo"
        End Select
                frmBusquedas.Show vbModal
                txtCodigoCorriente = gBusquedas.Resultado
                txtCodigoCorriente.SetFocus
        
        
    Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
        

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim i As Integer

Me.MousePointer = vbHourglass

With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Fecha = '" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
   .Formulas(1) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"

Select Case (MsgBox("Si Desea ver todos los códigos presione CANCELAR" & vbCrLf _
         & "Si desea ver solo los códigos Activos presione SI" & vbCrLf _
         & "Si desea ver solo los códigos Inactivos presione NO", vbYesNoCancel))
  Case vbYes
     .SelectionFormula = "{CATALOGO.CASOS} > 0"
  Case vbNo
     .SelectionFormula = "{CATALOGO.CASOS} = 0"
End Select

Select Case ButtonMenu.Key
    Case "OrdCodigo"
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDeCreditosResumen.rpt")
     .SortFields(0) = "+{Catalogo.Codigo}"
   
    
    Case "OrdDes"
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDeCreditosResumen.rpt")
     .SortFields(0) = "+{Catalogo.Descripcion}"
    
    Case "DetalladoCod"
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDeCreditosDetalle.rpt")
     .SortFields(0) = "+{Catalogo.Codigo}"
    
    Case "DetalladoDesc"
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDeCreditosDetalle.rpt")
     .SortFields(0) = "+{Catalogo.Descripcion}"

    Case "linea"
     If txtCodigoCorriente.Text <> "" Then
      .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDeCreditosDetalle.rpt")
      .SelectionFormula = "{Catalogo.codigo} = '" & txtCodigoCorriente.Text & "'"
    End If
    
    Case "repCuentasCod"
     .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoCodigoCuentas.rpt")
    
End Select
 
 .PrintReport

End With

Me.MousePointer = vbDefault
End Sub


Private Sub txtCodigoCorriente_Change()
Dim i As Integer

If txtCodigoCorriente <> "" Then
  mCodigoAlterno = ""
  For i = 1 To Len(Trim(txtCodigoCorriente))
     mCodigoAlterno = mCodigoAlterno & Mid(Trim(txtCodigoCorriente), Len(Trim(txtCodigoCorriente)) - (i - 1), 1)
  Next i
End If

End Sub

Private Sub txtCodigoCorriente_GotFocus()
 strConsulta = "codigo"
End Sub

Private Sub txtCodigoCorriente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then

        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Columna = "Codigo"
        gBusquedas.Orden = "Codigo"
        gBusquedas.Filtro = ""
        
        If chkFiltrarAutoGestion.Value = xtpChecked Then
            gBusquedas.Filtro = " and WEBSITE = 1"
        End If
        
        If chkFiltraActivas.Value = xtpChecked Then
            gBusquedas.Filtro = gBusquedas.Filtro & " and ACTIVO = " & chkFiltraActivas.Value
        End If
        
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
            txtCodigoCorriente = gBusquedas.Resultado
            txtCodigoCorriente_LostFocus
        End If
End If

End Sub

Private Sub sbFontsChecks(Chk As Object)

If Chk.Value = vbChecked Then
   Chk.ForeColor = vbBlue
Else
   Chk.ForeColor = vbBlack
End If

End Sub

Private Sub txtCodigoCorriente_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

txtCodigoCorriente.Text = UCase(txtCodigoCorriente.Text)

If Trim(txtCodigoCorriente.Text) <> "" And intEdita = 2 Then
  strSQL = "exec spCrd_Catalogo_Consulta'" & txtCodigoCorriente.Text & "'"
  Call OpenRecordSet(rs, strSQL)
Else
  If intEdita = 2 Then MsgBox "Digite la línea de crédito a consultar ...", vbInformation
  Exit Sub
End If


If Not rs.EOF And Not rs.BOF Then

 txtCodigoCorriente.Text = Trim(rs!Codigo)
 mCodigoAlterno = Trim(rs!codigoa)
 txtDescripcion.Text = Trim(rs!Descripcion)
  
 txtNotas.Text = Trim(rs!Notas)
  
 
 chkActivo.Value = rs!Activo
 chkLineaInterna.Value = rs!Linea_Interna
 chkCodigoAlterno.Value = rs!DEDUC_CODIGO_ALTER
 
 chkListaRefundibles.Value = rs!FILTRA_REFUNDIBLES
 
 
 chkPermite_PersonaEnCbrJud.Value = rs!Permite_PersonaEnCbrJud
 
 'Caracteristicas
 chkConvenio.Value = IIf((rs!Convenio = "S"), 1, 0)
 chkCodigoPoliza.Value = IIf((rs!Poliza = "S"), 1, 0)
 chkRefundeOtraOperacion.Value = IIf((rs!Refunde = "S"), 1, 0)
 chkRetencion.Value = IIf((rs!retencion = "S"), 1, 0)
 chkAceptaRefundicion.Value = IIf((rs!aceptarefun = "S"), 1, 0)
 chkPrimerCuota.Value = IIf((rs!PRIMER_CUOTA = "S"), 1, 0)
 chkPideCheque.Value = IIf((rs!pideCheque = "S"), 1, 0)
 
 chkRetencionSaldo.Value = rs!RETENCION_MUESTRA_SALDO
 
 chkCobertura.Value = rs!Cobertura
 chkLineaMora.Value = IIf(IsNull(rs!GENERA_MORA), 0, rs!GENERA_MORA)
 
 chkAceptaMovCajas.Value = rs!MovCajas
  
 
 
 'Carga tramites

 If rs!liq_tipoAumento = "P" Then
  cboLiqTasa.Text = "Aumento de Puntos"
 Else
  cboLiqTasa.Text = "Tasa Fija"
 End If
 txtLiqValor.Text = CStr(rs!Liq_Valor)
 
 
 cboFactorCalculo.Text = fxCrd_Factor_Calculo(rs!Base_Calculo)
 
 If rs!COBRO_TIPO_APLICACION = "H" Then
    cboMetodoCancelaCuotas.Text = "Horizontal"
 Else
    cboMetodoCancelaCuotas.Text = "Vertical"
 End If
 
 
 If rs!Tramite = "C" Then
    cboTipoCredito.Text = "Credito"
 Else
    cboTipoCredito.Text = "Microcredito"
 End If
 
 
 If rs!Requisitos_Tipo = "L" Then
    cboRequisitos.Text = "Línea"
 Else
    cboRequisitos.Text = "Garantía"
 End If
 
 'Parametros
 cmdTabla.Enabled = True
 cmdCuentas.Enabled = True
 
 Call sbCboAsignaDato(cboComite, Trim(rs!ComiteDesc), True, rs!id_Comite)
 Call sbCboAsignaDato(cboInstitucion, Trim(rs!InstitucionDesc), True, rs!cod_institucion)
 Call sbCboAsignaDato(cboMoneda, rs!DivisaDesc & "", True, Trim(rs!DivisaId & ""))
 
 txtMembresia.Text = CStr(rs!Membresia_Meses)
 txtDias.Text = CStr(rs!TramiteDias)
 txtNumOperaciones.Text = CStr(rs!operaciones_activas)
 
 
 txtPorcRefun.Text = CStr(rs!refunde_porc)
 
 cboTipoRefun.Text = IIf((rs!refunde_tipo = "P"), "01 - Plazo", "02 - Monto")
 
 txtPorcCancelacion.Text = CStr(rs!PORC_CARGO_CANCELACION)
 txtAnticipoMesesPenalizados.Text = CStr(rs!ANTICIPO_MESES)
 
 If Not IsNull(rs!FechaCorte) Then
   dtpFechaCorte = rs!FechaCorte
 Else
   dtpFechaCorte = rs!FechaServer
 End If
 
 chkFechaCorte.Value = IIf((rs!FechaCorteAlterna = "S"), 1, 0)
 If chkFechaCorte.Value = vbChecked Then
   dtpFechaCorte.Enabled = True
 Else
   dtpFechaCorte.Enabled = False
 End If
 
 chkRngTasaDestino.Value = rs!Tasa_Destino
 chkRngTBP.Value = rs!TBP_Utiliza
 txtRngPuntosAdicionalesTBP.Text = CStr(rs!TBP_Adicional)
 
 Select Case rs!Tasa_Mora_Tipo
    Case "POR"
        cboMoraTipo.Text = "Porcentaje s/Tasa Vigente"
    Case "PTS"
       cboMoraTipo.Text = "Puntos Adicionales"
    Case "N/A"
       cboMoraTipo.Text = "No Calcula Int.Moratorio"
    Case "TF"
       cboMoraTipo.Text = "Tasa Fija"
 End Select
 
 txtTasaMora.Text = CStr(rs!tasa_mora_add)


 'Revolutivo
 
 chkRevLinea.Value = rs!Revolutiva
 chkRevTopeRetiros.Value = rs!Revolutiva_Tope_Retiros
 chkRevEstudio.Value = rs!Revolutiva_Estudio
 chkRevPlanAhorros.Value = rs!Revolutiva_Plan_Ahorro_Utiliza
 
 txtRevPlanAhorro.Text = rs!Revolutiva_Plan_Ahorro & ""
 txtRevPlanDesc.Text = rs!PlanAhorroDesc & ""
  
 'Reservas
 chkReserva_Aplica.Value = rs!Reserva_Aplica
 chkReserva_Flat.Value = rs!Reserva_Facial_Flat
 chkReserva_Mora.Value = rs!Reserva_Mora_Apl
 txtReserva_Plan.Text = rs!Reserva_Codigo & ""
 txtReserva_PlanDesc.Text = rs!Reserva_PlanDesc & ""
 txtReserva_MontoMin.Text = Format(rs!Reserva_Monto_Minimo, "Standard")

 'Oficinas
 chkOficinaLinea.Value = rs!Oficina_Linea
 txtOficina.Text = rs!Oficina & ""
 txtOficinaDesc.Text = rs!OficinaDesc & ""
 
 'Auto Gestion
 chkLineaVisibleEC.Value = IIf(IsNull(rs!visible_EC), 0, rs!visible_EC)
 chkWebSite.Value = rs!WebSite
 
 chkFP_POS.Value = rs!FORMA_PAGO_POS
 chkFP_Web.Value = rs!FORMA_PAGO_WEB
 
 chkGirosPorLinea.Value = rs!AUTO_GESTION_LMAX
 txtGiroMaxTransac.Text = Format(rs!GIRO_MAX_TRANSAC, "Standard")
 
 If rs!AUTO_GESTION_TIPO = "C" Then
    rbTipoCrdWeb.Item(0).Value = True
 Else
    rbTipoCrdWeb.Item(1).Value = True
 End If
 
 chkGirosBancos.Value = rs!GIRO_AUTOMATICO
 txtGirosMntTraslado.Text = Format(rs!GIRO_MONTO_BASE, "Standard")
 
 chkRefundeAuto.Value = rs!REFUNDE_AUTO
 chkRefundeAumentaBase.Value = rs!REFUNDE_AUMENTA_BASE
 
 txtGiroMinimo.Text = Format(rs!GIRO_MINIMO, "Standard")
 
 
 '2024-03-01 Nuevos
 
 chkNotifica_Formaliza.Value = rs!IND_NOTIFICA_CLI_FORMALIZA
 chkNotifica_Cancela.Value = rs!IND_NOTIFICA_CLI_CANCELA
 
 chkBonifica.Value = rs!IND_MOV_APLICA_BONIF
 chkPago_Activa.Value = rs!IND_PAGO_OP_APLICACION
 chkReadecua.Value = rs!IND_READECUA
 chkMntMax.Value = rs!IND_MONTO_MAX
 
 chkSupervision.Value = rs!ID_REQ_SUPERVISION
 txtSupervisionMonto.Text = Format(rs!MONTO_SUPERVISION, "Standard")
 
 chkSINPE.Value = rs!MOV_SINPE
 
 Select Case rs!MOV_SINPE_TIPOS
    Case 1
         cboSinpe.Text = "Trámite Interbancario"
    Case 3
         cboSinpe.Text = "No agregar motivo"
 End Select
 
 
 chkTasaFija_TBP_Apl.Value = rs!TASA_FIJA_X_TBP
 txtTasaFija_TBP_Pts.Text = Format(rs!TASA_FIJA_X_TBP_PUNTOS_ADD, "Standard")
 txtTasaFija_Plazo.Text = CStr(rs!PLAZO_TASA_FIJA)
 
 chkEdadPension_Estudio.Value = rs!IND_EDAD_PENSION_EST
 chkEdadPension_Formalizacion.Value = rs!IND_EDAD_PENSION_FOR
 
 txtAnticipo_Extraordinario_Porc.Text = Format(rs!PORC_ANTICIPO_EXT, "Standard")
 
 
 'Activa Tabs
 For i = 0 To tcMain.ItemCount - 1
   tcMain.Item(i).Enabled = True
 Next i
 
 tcMain.Item(0).Selected = True
 tcRangos.Item(0).Selected = True
 tcParametros.Item(0).Selected = True
 
 Call sbToolBar(tlbPrincipal, "activo")
 
 Call sbCargaEstados
 Call sbCargaGarantias
 
Else
 
 MsgBox "Código no existe en la Base de Datos", vbOKOnly
 
 lswEstados.Visible = False
 lswGarantias.Visible = False


End If

rs.Close
 
Call RefrescaTags(Me)

vGrid.Enabled = cmdTabla.Enabled
 
 lswGarantias.ShowBorder = True
 lswEstados.ShowBorder = True
 
vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 vPaso = False

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer
 
vPaso = True
 
txtCodigoCorriente.Text = ""
mCodigoAlterno = ""
txtDescripcion.Text = ""
 
chkPideCheque.Value = xtpUnchecked


chkActivo.Value = xtpChecked
chkLineaInterna.Value = xtpChecked
chkCobertura.Value = xtpUnchecked

chkPermite_PersonaEnCbrJud.Value = xtpUnchecked

chkRefundeOtraOperacion.Value = xtpChecked

chkListaRefundibles.Value = xtpChecked
chkConvenio.Value = xtpUnchecked
chkCodigoPoliza.Value = xtpUnchecked
chkRetencion = xtpUnchecked
chkRetencionSaldo.Value = xtpChecked
chkAceptaRefundicion.Value = xtpChecked
chkPrimerCuota.Value = xtpUnchecked


txtNotas.Text = ""

chkFechaCorte.Value = xtpUnchecked
dtpFechaCorte.Enabled = False

txtMembresia.Text = 0
txtDias.Text = 30

txtNumOperaciones.Text = 1


cboFactorCalculo.Text = "Comercial Nivelada (30/360)"


txtPorcRefun.Text = 25
cboTipoRefun.Text = "01 - Plazo"

txtPorcCancelacion.Text = 0
txtAnticipoMesesPenalizados.Text = 12

chkLineaVisibleEC.Value = xtpChecked
chkLineaMora.Value = xtpChecked
chkAceptaMovCajas.Value = xtpChecked


 chkNotifica_Formaliza.Value = xtpUnchecked
 chkNotifica_Cancela.Value = xtpUnchecked

 chkBonifica.Value = xtpUnchecked
 chkPago_Activa.Value = xtpUnchecked
 chkReadecua.Value = xtpUnchecked
 chkMntMax.Value = xtpUnchecked
 chkSupervision.Value = xtpUnchecked
 txtSupervisionMonto.Text = "0"
 chkSINPE.Value = xtpUnchecked
 cboSinpe.Text = "No agregar motivo"
 cboTipoCredito.Text = "Credito"

 chkTasaFija_TBP_Apl.Value = xtpUnchecked
 txtTasaFija_TBP_Pts.Text = "0"
 txtTasaFija_Plazo.Text = "0"
 
 chkEdadPension_Estudio.Value = xtpChecked
 chkEdadPension_Formalizacion.Value = xtpChecked
 txtAnticipo_Extraordinario_Porc.Text = "0"




lswCuentas.ListItems.Clear

lswGarantias.ListItems.Clear
lswEstados.ListItems.Clear

lswGarantias.Visible = False
lswEstados.Visible = False

cmdTabla.Enabled = False
cmdCuentas.Enabled = False

tcMain.Item(0).Selected = True
For i = xtpUnchecked To tcMain.ItemCount - 1
   tcMain.Item(i).Selected = False
Next i

tcRangos.Item(0).Selected = True
tcParametros.Item(0).Selected = True

 
chkRngTasaDestino.Value = xtpUnchecked
chkRngTBP.Value = xtpUnchecked

txtRngPuntosAdicionalesTBP.Text = 0
txtTasaMora.Text = 3

 chkRevLinea.Value = xtpUnchecked
 chkRevTopeRetiros.Value = xtpUnchecked
 chkRevEstudio.Value = xtpUnchecked
 chkRevPlanAhorros.Value = xtpUnchecked
 
 txtRevPlanAhorro.Text = ""
 txtRevPlanDesc.Text = ""

 chkReserva_Aplica.Value = xtpUnchecked
 chkReserva_Flat.Value = xtpUnchecked
 chkReserva_Mora.Value = xtpUnchecked
 txtReserva_Plan.Text = ""
 txtReserva_PlanDesc.Text = ""
 txtReserva_MontoMin.Text = "0"
 
 chkOficinaLinea.Value = xtpUnchecked
 txtOficina.Text = ""
 txtOficinaDesc.Text = ""


chkWebSite.Value = xtpUnchecked
chkFP_POS.Value = xtpUnchecked
chkFP_Web.Value = xtpUnchecked
chkGirosPorLinea.Value = xtpUnchecked
rbTipoCrdWeb.Item(0).Value = True


chkGirosBancos.Value = xtpUnchecked
txtGiroMinimo.Text = "0"
txtGiroMaxTransac.Text = "0"
txtGirosMntTraslado.Text = "0"
             
chkRefundeAuto.Value = xtpUnchecked
chkRefundeAumentaBase.Value = xtpUnchecked

vPaso = False
 
End Sub


Private Function fxPrioridad(vTipoGarantia As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDefecto As Integer

strSQL = "select isnull(max(prioridad),0) as Prioridad from catalogo where prioridad between "

Select Case UCase(vTipoGarantia)
  Case "P" 'Polizas
    strSQL = strSQL & "0 and 999"
    vDefecto = 1
  Case "A" 'sobre ahorros
    strSQL = strSQL & "1000 and 1999"
    vDefecto = 1000
  Case "F", "X" 'Fiduciarios y Prendas / Acciones
    strSQL = strSQL & "2000 and 2999"
    vDefecto = 2000
  Case "H" 'Vivienda
    strSQL = strSQL & "3000 and 3999"
    vDefecto = 3000
  Case "S" 'sin Garantias
    strSQL = strSQL & "4000 and 4999"
    vDefecto = 4000
  Case "R" 'Retenciones
    strSQL = strSQL & "5000 and 5999"
    vDefecto = 5000
  Case Else
      vDefecto = 6000
End Select

Call OpenRecordSet(rs, strSQL)
If rs!prioridad = 0 Then
 fxPrioridad = vDefecto
Else
 fxPrioridad = rs!prioridad + 1
End If
rs.Close

End Function


Private Sub sbGuardaLinea()
Dim strSQL As String, strTramite As String
Dim vMoraTipo As String, vBaseCalculo As String, vGestionTipo As String

On Error GoTo vError

txtDescripcion.Text = UCase(txtDescripcion.Text)
txtCodigoCorriente.Text = UCase(txtCodigoCorriente.Text)
mCodigoAlterno = UCase(mCodigoAlterno)

vBaseCalculo = fxCrd_Factor_Calculo(cboFactorCalculo.Text)

Select Case True
    Case rbTipoCrdWeb.Item(0).Value
      vGestionTipo = "C"
    Case rbTipoCrdWeb.Item(1).Value
      vGestionTipo = "S"
End Select
    
strTramite = cboTipoCredito.ItemData(cboTipoCredito.ListIndex)

Select Case cboMoraTipo.Text
 Case "Puntos Adicionales"
    vMoraTipo = "PTS"
 Case "Porcentaje s/Tasa Vigente"
    vMoraTipo = "POR"
 Case "No Calcula Int.Moratorio"
    vMoraTipo = "N/A"
 Case "Tasa Fija"
    vMoraTipo = "TF"
End Select


strSQL = "exec spCrd_Catalogo_Registro '" & txtCodigoCorriente.Text & "','" & mCodigoAlterno & "','" & txtDescripcion.Text _
           & "'," & cboComite.ItemData(cboComite.ListIndex) & ",'" & strTramite & "',0,'" & IIf((chkConvenio.Value = 1), "S", "N") _
           & "','" & IIf((chkRefundeOtraOperacion.Value = 1), "S", "N") & "','" & IIf((chkCodigoPoliza.Value = 1), "S", "N") _
           & "','" & IIf((chkRetencion.Value = 1), "S", "N") & "','" & IIf((chkAceptaRefundicion.Value = 1), "S", "N") _
           & "','" & cboMoneda.ItemData(cboMoneda.ListIndex) & "'," & txtDias & ",'" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") _
           & "','" & IIf((chkFechaCorte.Value = 1), "S", "N") & "','" & IIf(Mid(cboTipoRefun.Text, 1, 2) = "01", "P", "M") _
           & "','" & txtPorcRefun.Text & "'," & txtNumOperaciones.Text & "," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",'" & IIf((chkPrimerCuota.Value = 1), "S", "N") _
           & "','" & IIf((chkPideCheque.Value = 1), "S", "N") & "'," & chkActivo.Value & ",'" & IIf(Mid(cboLiqTasa.Text, 1, 1) = "A", "P", "F") _
           & "'," & txtLiqValor.Text & "," & chkRngTBP.Value & "," & txtRngPuntosAdicionalesTBP.Text & "," & chkRngTasaDestino.Value _
           & "," & CCur(txtTasaMora) & ",'" & vMoraTipo & "'," & chkLineaInterna.Value & "," & chkLineaVisibleEC.Value & "," & chkLineaMora.Value _
           & "," & chkAceptaMovCajas.Value & "," & chkWebSite.Value & ",'" & txtNotas.Text & "'," & chkCobertura.Value _
           & "," & txtPorcCancelacion.Text & ",'" & Mid(cboRequisitos.Text, 1, 1) & "','" & vBaseCalculo _
           & "','" & Mid(cboMetodoCancelaCuotas.Text, 1, 1) & "'," & txtAnticipoMesesPenalizados.Text & "," & chkRetencionSaldo.Value _
           & "," & chkRevLinea.Value & "," & chkRevTopeRetiros.Value & "," & chkRevEstudio.Value & "," & chkRevPlanAhorros.Value _
           & ",'" & txtRevPlanAhorro.Text & "'," & chkCodigoAlterno.Value & "," & chkPermite_PersonaEnCbrJud.Value _
           & "," & chkReserva_Aplica.Value & "," & chkReserva_Flat.Value & "," & chkReserva_Mora.Value & ",'" & txtReserva_Plan.Text _
           & "'," & CCur(txtReserva_MontoMin.Text) & "," & chkOficinaLinea.Value & "," & IIf((Trim(txtOficina.Text) = ""), "Null", "'" & txtOficina.Text & "'") _
           & "," & chkFP_POS.Value & "," & chkFP_Web.Value & "," & chkListaRefundibles.Value & "," & txtMembresia.Text _
           & ",'" & vGestionTipo & "'," & chkGirosPorLinea.Value & ", " & chkGirosBancos.Value & ", " & CCur(txtGirosMntTraslado.Text) & ", " & CCur(txtGiroMinimo.Text) _
           & ", " & CCur(txtGiroMaxTransac.Text) & ", " & chkRefundeAuto.Value & ", " & chkRefundeAumentaBase.Value _
           & ", " & chkSINPE.Value & ", " & cboSinpe.ItemData(cboSinpe.ListIndex) & ", " & chkTasaFija_TBP_Apl.Value & ", " & CCur(txtTasaFija_TBP_Pts.Text) _
           & ", " & txtTasaFija_Plazo.Text & ", " & chkSupervision.Value & ", " & CCur(txtSupervisionMonto.Text) & ", " & CCur(txtAnticipo_Extraordinario_Porc.Text) _
           & ", " & chkBonifica.Value & ", " & chkPago_Activa.Value & ", " & chkReadecua.Value & ", " & chkMntMax.Value & ", " & chkEdadPension_Estudio.Value & ", " & chkEdadPension_Formalizacion.Value _
           & ", " & chkNotifica_Formaliza.Value & ", " & chkNotifica_Cancela.Value & ", '" & glogon.Usuario & "', 'A'"
Call ConectionExecute(strSQL)
                


Select Case intEdita
 Case 0 'Inserta
 
 
'
'    iPrioridad = fxPrioridad("A")
'    If chkRetencion.Value = 1 Then iPrioridad = fxPrioridad("R")
'    If chkCodigoPoliza.Value = 1 Then iPrioridad = fxPrioridad("P")
'
'    strSQL = "insert into catalogo(codigo, codigoa, descripcion, id_comite, tramite" _
'           & ", premio, convenio, refunde, poliza, retencion, AceptaRefun, casos, prioridad, moneda, TramiteDias" _
'           & ", fechaCorte, FechaCorteAlterna, refunde_tipo, refunde_porc, operaciones_activas" _
'           & ", cod_institucion, primer_cuota, PideCheque, Activo, Liq_TipoAumento, Liq_Valor, TBP_Utiliza, TBP_Adicional, Tasa_Destino, TASA_MORA_ADD, Tasa_Mora_Tipo" _
'           & ", Linea_Interna, Visible_EC, Genera_Mora, MovCajas, WebSite, Notas, Cobertura, PORC_CARGO_CANCELACION, Requisitos_Tipo" _
'           & ", Base_Calculo,COBRO_TIPO_APLICACION,ANTICIPO_MESES,RETENCION_MUESTRA_SALDO" _
'           & ", Revolutiva,Revolutiva_Tope_Retiros,Revolutiva_Estudio,Revolutiva_Plan_Ahorro_Utiliza,Revolutiva_Plan_Ahorro,DEDUC_CODIGO_ALTER,Permite_PersonaEnCbrJud" _
'           & ", Reserva_Aplica, Reserva_Facial_Flat, Reserva_Mora_Apl, Reserva_Codigo , Reserva_Monto_Minimo, Oficina_Linea,Oficina_Codigo" _
'           & ", FORMA_PAGO_POS, FORMA_PAGO_WEB, FILTRA_REFUNDIBLES, MEMBRESIA_MESES, AUTO_GESTION_TIPO, AUTO_GESTION_LMAX" _
'           & ", GIRO_AUTOMATICO, GIRO_MONTO_BASE, GIRO_MINIMO, GIRO_MAX_TRANSAC, REFUNDE_AUTO, REFUNDE_AUMENTA_BASE, IMPUESTO" _
'           & ", MOV_SINPE, MOV_SINPE_TIPOS, TASA_FIJA_X_TBP, TASA_FIJA_X_TBP_PUNTOS_ADD, PLAZO_TASA_FIJA, ID_REQ_SUPERVISION, MONTO_SUPERVISION, PORC_ANTICIPO_EXT" _
'           & ", IND_MOV_APLICA_BONIF, IND_PAGO_OP_APLICACION, IND_READECUA, IND_MONTO_MAX, IND_EDAD_PENSION_EST, IND_EDAD_PENSION_FOR" _
'           & ", IND_NOTIFICA_CLI_FORMALIZA, IND_NOTIFICA_CLI_CANCELA, REGISTRO_FECHA, REGISTRO_USUARIO)"
'
'
'    strSQL = strSQL & " values('" & txtCodigoCorriente.Text & "','" & mCodigoAlterno & "','" & txtDescripcion _
'           & "'," & cboComite.ItemData(cboComite.ListIndex) & ",'" & strTramite & "',0,'" & IIf((chkConvenio.Value = 1), "S", "N") _
'           & "','" & IIf((chkRefundeOtraOperacion.Value = 1), "S", "N") & "','" & IIf((chkCodigoPoliza.Value = 1), "S", "N") _
'           & "','" & IIf((chkRetencion.Value = 1), "S", "N") & "','" & IIf((chkAceptaRefundicion.Value = 1), "S", "N") _
'           & "',0," & iPrioridad & ",'" & cboMoneda.ItemData(cboMoneda.ListIndex) & "'," & txtDias & ",'" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") _
'           & "','" & IIf((chkFechaCorte.Value = 1), "S", "N") & "','" & IIf(Mid(cboTipoRefun.Text, 1, 2) = "01", "P", "M") _
'           & "','" & txtPorcRefun.Text & "'," & txtNumOperaciones & "," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",'" & IIf((chkPrimerCuota.Value = 1), "S", "N") _
'           & "','" & IIf((chkPideCheque.Value = 1), "S", "N") & "'," & chkActivo.Value & ",'" & IIf(Mid(cboLiqTasa.Text, 1, 1) = "A", "P", "F") _
'           & "'," & txtLiqValor.Text & "," & chkRngTBP.Value & "," & txtRngPuntosAdicionalesTBP.Text & "," & chkRngTasaDestino.Value _
'           & "," & CCur(txtTasaMora) & ",'" & vMoraTipo & "'," & chkLineaInterna.Value & "," & chkLineaVisibleEC.Value & "," & chkLineaMora.Value _
'           & "," & chkAceptaMovCajas.Value & "," & chkWebSite.Value & ",'" & txtNotas.Text & "'," & chkCobertura.Value _
'           & "," & txtPorcCancelacion.Text & ",'" & Mid(cboRequisitos.Text, 1, 1) & "','" & vBaseCalculo _
'           & "','" & Mid(cboMetodoCancelaCuotas.Text, 1, 1) & "'," & txtAnticipoMesesPenalizados.Text & "," & chkRetencionSaldo.Value _
'           & "," & chkRevLinea.Value & "," & chkRevTopeRetiros.Value & "," & chkRevEstudio.Value & "," & chkRevPlanAhorros.Value _
'           & ",'" & txtRevPlanAhorro.Text & "'," & chkCodigoAlterno.Value & "," & chkPermite_PersonaEnCbrJud.Value _
'           & "," & chkReserva_Aplica.Value & "," & chkReserva_Flat.Value & "," & chkReserva_Mora.Value & ",'" & txtReserva_Plan.Text
'
'
'     strSQL = strSQL & "'," & CCur(txtReserva_MontoMin.Text) & "," & chkOficinaLinea.Value & "," & IIf((Trim(txtOficina.Text) = ""), "Null", "'" & txtOficina.Text & "'") _
'            & "," & chkFP_POS.Value & "," & chkFP_Web.Value & "," & chkListaRefundibles.Value & "," & txtMembresia.Text _
'            & ",'" & vGestionTipo & "'," & chkGirosPorLinea.Value & ", " & chkGirosBancos.Value & ", " & CCur(txtGirosMntTraslado.Text) & ", " & CCur(txtGiroMinimo.Text) _
'            & ", " & CCur(txtGiroMaxTransac.Text) & ", " & chkRefundeAuto.Value & ", " & chkRefundeAumentaBase.Value & ", 0" _
'            & ", " & chkSINPE.Value & ", " & cboSinpe.ItemData(cboSinpe.ListIndex) & ", " & chkTasaFija_TBP_Apl.Value & ", " & CCur(txtTasaFija_TBP_Pts.Text) _
'            & ", " & txtTasaFija_Plazo.Text & ", " & chkSupervision.Value & ", " & CCur(txtSupervisionMonto.Text) & ", " & CCur(txtAnticipo_Extraordinario_Porc.Text) _
'            & ", " & chkBonifica.Value & ", " & chkPago_Activa.Value & ", " & chkReadecua.Value & ", " & chkMntMax.Value & ", " & chkEdadPension_Estudio.Value & ", " & chkEdadPension_Formalizacion.Value _
'            & ", " & chkNotifica_Formaliza.Value & ", " & chkNotifica_Cancela.Value & ", getdate(), '" & glogon.Usuario & "')"
'
'    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", "Linea de Credito : " & Trim(txtCodigoCorriente.Text))
  
    Call sbCargaEstados
    Call sbCargaGarantias
  
  Case 1 'Edita
'    strSQL = "update catalogo set descripcion = '" & txtDescripcion.Text & "', id_comite = " & cboComite.ItemData(cboComite.ListIndex) & ", tramite = '" & strTramite & "'" _
'             & ", convenio = '" & IIf((chkConvenio.Value = 1), "S", "N") & "', DEDUC_CODIGO_ALTER = " & chkCodigoAlterno.Value _
'             & ", poliza = '" & IIf((chkCodigoPoliza.Value = 1), "S", "N") & "'" _
'             & ", refunde = '" & IIf((chkRefundeOtraOperacion.Value = 1), "S", "N") & "'" _
'             & ", Aceptarefun = '" & IIf((chkAceptaRefundicion.Value = 1), "S", "N") & "'" _
'             & ", retencion = '" & IIf((chkRetencion.Value = 1), "S", "N") & "', RETENCION_MUESTRA_SALDO = " & chkRetencionSaldo.Value _
'             & ", primer_cuota = '" & IIf((chkPrimerCuota.Value = 1), "S", "N") & "'" _
'             & ", fechaCorteAlterna = '" & IIf((chkFechaCorte.Value = 1), "S", "N") & "'" _
'             & ", fechaCorte = '" & Format(dtpFechaCorte, "yyyy/mm/dd") & "'" _
'             & ", moneda = '" & cboMoneda.ItemData(cboMoneda.ListIndex) & "', refunde_tipo = '" & IIf(Mid(cboTipoRefun, 1, 2) = "01", "P", "M") & "'" _
'             & ", refunde_porc = " & txtPorcRefun & ", codigoa = '" & Trim(mCodigoAlterno) & "'" _
'             & ", TramiteDias = " & txtDias & ", Operaciones_activas = " & txtNumOperaciones _
'             & ", cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ", PideCheque = '" & IIf((chkPideCheque.Value = 1), "S", "N") & "'" _
'             & ",Activo = " & chkActivo.Value & ",LIQ_TIPOAUMENTO = '" & IIf(Mid(cboLiqTasa.Text, 1, 1) = "A", "P", "F") _
'             & "',Liq_Valor = " & txtLiqValor & ",TBP_Utiliza = " & chkRngTBP.Value & ",TBP_Adicional = " & txtRngPuntosAdicionalesTBP.Text _
'             & ",Tasa_Destino = " & chkRngTasaDestino.Value & ",TASA_MORA_ADD = " & CCur(txtTasaMora) & ", Tasa_Mora_Tipo ='" & vMoraTipo _
'             & "', Linea_Interna = " & chkLineaInterna.Value & ", Requisitos_Tipo = '" & Mid(cboRequisitos.Text, 1, 1) & "'"
'
'     strSQL = strSQL & ",Visible_EC = " & chkLineaVisibleEC.Value & ", Genera_Mora = " & chkLineaMora.Value & ", cobertura = " & chkCobertura.Value _
'             & ", MovCajas = " & chkAceptaMovCajas.Value & ", WebSite = " & chkWebSite.Value & ", Notas = '" & Trim(txtNotas.Text) _
'             & "',PORC_CARGO_CANCELACION = " & txtPorcCancelacion.Text & ",Base_Calculo = '" & vBaseCalculo _
'             & "', COBRO_TIPO_APLICACION = '" & Mid(cboMetodoCancelaCuotas.Text, 1, 1) & "', ANTICIPO_MESES = " & txtAnticipoMesesPenalizados.Text _
'             & ", Revolutiva = " & chkRevLinea.Value & ", Revolutiva_Tope_Retiros = " & chkRevTopeRetiros.Value _
'             & ", Revolutiva_Estudio = " & chkRevEstudio.Value & ", Revolutiva_Plan_Ahorro_Utiliza = " & chkRevPlanAhorros.Value _
'             & ", Revolutiva_Plan_Ahorro = '" & Trim(txtRevPlanAhorro.Text) & "'" _
'             & ", Permite_PersonaEnCbrJud = " & chkPermite_PersonaEnCbrJud.Value _
'             & ", Reserva_Aplica = " & chkReserva_Aplica.Value & ", Reserva_Facial_Flat = " & chkReserva_Flat.Value _
'             & ", Reserva_Mora_Apl = " & chkReserva_Mora.Value & ", Reserva_Codigo = '" & txtReserva_Plan.Text _
'             & "', Reserva_Monto_Minimo = " & CCur(txtReserva_MontoMin.Text) _
'             & ", Oficina_Linea = " & chkOficinaLinea.Value & ", Oficina_Codigo = " & IIf((Trim(txtOficina.Text) = ""), "Null", "'" & txtOficina.Text & "'") _
'             & ", FORMA_PAGO_POS = " & chkFP_POS.Value & ", FORMA_PAGO_WEB = " & chkFP_Web.Value & ", FILTRA_REFUNDIBLES = " & chkListaRefundibles.Value _
'             & ", Membresia_Meses = " & txtMembresia.Text _
'             & ", AUTO_GESTION_TIPO = '" & vGestionTipo & "', AUTO_GESTION_LMAX = " & chkGirosPorLinea.Value _
'             & ", GIRO_AUTOMATICO = " & chkGirosBancos.Value & ", GIRO_MONTO_BASE = " & CCur(txtGirosMntTraslado.Text) & ",  GIRO_MINIMO = " & CCur(txtGiroMinimo.Text) _
'             & ", GIRO_MAX_TRANSAC = " & CCur(txtGiroMaxTransac.Text) & ", REFUNDE_AUTO = " & chkRefundeAuto.Value & ", REFUNDE_AUMENTA_BASE = " & chkRefundeAumentaBase.Value _
'
'     strSQL = strSQL & ", MOV_SINPE = " & chkSINPE.Value & ", MOV_SINPE_TIPOS = " & cboSinpe.ItemData(cboSinpe.ListIndex) _
'             & ", TASA_FIJA_X_TBP = " & chkTasaFija_TBP_Apl.Value & ", TASA_FIJA_X_TBP_PUNTOS_ADD = " & CCur(txtTasaFija_TBP_Pts.Text) & ", PLAZO_TASA_FIJA = " & txtTasaFija_Plazo.Text _
'             & ", ID_REQ_SUPERVISION = " & chkSupervision.Value & ", MONTO_SUPERVISION = " & CCur(txtSupervisionMonto.Text) & ", PORC_ANTICIPO_EXT = " & CCur(txtAnticipo_Extraordinario_Porc.Text) _
'             & ", IND_MOV_APLICA_BONIF = " & chkBonifica.Value & ", IND_PAGO_OP_APLICACION = " & chkPago_Activa.Value & ", IND_READECUA = " & chkReadecua.Value _
'             & ", IND_MONTO_MAX = " & chkMntMax.Value & ", IND_EDAD_PENSION_EST = " & chkEdadPension_Estudio.Value & ", IND_EDAD_PENSION_FOR = " & chkEdadPension_Formalizacion.Value _
'             & ", IND_NOTIFICA_CLI_FORMALIZA = " & chkNotifica_Formaliza.Value & ", IND_NOTIFICA_CLI_CANCELA = " & chkNotifica_Cancela.Value _
'             & ", MODIFICA_FECHA = getdate(), REGISTRO_USUARIO = '" & glogon.Usuario & "'" _
'             & " where codigo = '" & txtCodigoCorriente.Text & "'"
'    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Linea de Credito : " & Trim(txtCodigoCorriente))
    
End Select

intEdita = 2

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub txtDescripcion_GotFocus()
 strConsulta = "descripcion"
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then

        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Filtro = ""
        
        If chkFiltrarAutoGestion.Value = xtpChecked Then
            gBusquedas.Filtro = " and WEBSITE = 1"
        End If
        
        If chkFiltraActivas.Value = xtpChecked Then
            gBusquedas.Filtro = gBusquedas.Filtro & " and ACTIVO = " & chkFiltraActivas.Value
        End If
        
       
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
            txtCodigoCorriente = gBusquedas.Resultado
            txtCodigoCorriente_LostFocus
        End If
End If

End Sub


Private Function fxGuardarRango() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarRango = 0
vGrid.Sheet = 1
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
If vGrid.Text = "" Or vGrid.Text = "0" Then
   vGrid.col = 2
   strSQL = "insert into rangos(codigo,de,hasta,plazo,intc_soc,intm_soc,intc_nsoc,intm_nsoc)" _
          & " values('" & txtCodigoCorriente & "'," & CCur(vGrid.Text) & ","
   vGrid.col = 3
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 4
   strSQL = strSQL & vGrid.Text & ","
   vGrid.col = 5
   strSQL = strSQL & vGrid.Text & ","
   vGrid.col = 6
   strSQL = strSQL & vGrid.Text & ","
   vGrid.col = 7
   strSQL = strSQL & vGrid.Text & ","
   vGrid.col = 8
   strSQL = strSQL & vGrid.Text & ")"
     
   Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(max(consec),0) as ultimo from rangos where codigo = '" _
           & txtCodigoCorriente & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
    Call Bitacora("Registra", "Rango para el Codigo: " & txtCodigoCorriente & " ID:" & vGrid.Text)
   
   Else 'Actualizar
    vGrid.col = 2
    strSQL = "update rangos set de = " & CCur(vGrid.Text)
    vGrid.col = 3
    strSQL = strSQL & ",hasta = " & CCur(vGrid.Text)
    vGrid.col = 4
    strSQL = strSQL & ",plazo = " & vGrid.Text
    vGrid.col = 5
    strSQL = strSQL & ",intc_soc = " & vGrid.Text
    vGrid.col = 6
    strSQL = strSQL & ",intm_soc = " & vGrid.Text
    vGrid.col = 7
    strSQL = strSQL & ",intc_nsoc = " & vGrid.Text
    vGrid.col = 8
    strSQL = strSQL & ",intm_nsoc = " & vGrid.Text
    vGrid.col = 1
    strSQL = strSQL & " where consec = " & vGrid.Text
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Rango para el Codigo: " & txtCodigoCorriente & " ID:" & vGrid.Text)
    
   End If

   vGrid.col = 1
   fxGuardarRango = vGrid.Text
   
   Exit Function
   
vError:
 fxGuardarRango = 0
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function


Private Function fxGuardarRangoPlazo() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarRangoPlazo = 0

vGrid.Sheet = 2
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
If vGrid.Text = "" Or vGrid.Text = "0" Then
   vGrid.col = 2
   
   strSQL = "insert into rangos_Plazo(codigo,desde,hasta,tasa)" _
          & " values('" & txtCodigoCorriente & "'," & CCur(vGrid.Text) & ","
   vGrid.col = 3
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 4
   strSQL = strSQL & vGrid.Text & ")"
   
   Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(max(consec),0) as ultimo from rangos_Plazo where codigo = '" _
           & txtCodigoCorriente & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
    Call Bitacora("Registra", "Rango Plazo para el Codigo: " & txtCodigoCorriente & " ID:" & vGrid.Text)
   
   Else 'Actualizar
    vGrid.col = 2
    strSQL = "update rangos_Plazo set desde = " & CCur(vGrid.Text)
    vGrid.col = 3
    strSQL = strSQL & ",hasta = " & CCur(vGrid.Text)
    vGrid.col = 4
    strSQL = strSQL & ",Tasa = " & CCur(vGrid.Text)
    vGrid.col = 1
    strSQL = strSQL & " where consec = " & vGrid.Text
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Rango Plazo para el Codigo: " & txtCodigoCorriente & " ID:" & vGrid.Text)
    
   End If

   vGrid.col = 1
   fxGuardarRangoPlazo = vGrid.Text
   
   Exit Function
   
vError:
 fxGuardarRangoPlazo = 0
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function





Private Sub txtGiroMaxTransac_GotFocus()
On Error GoTo vError
  txtGiroMaxTransac.Text = CCur(txtGiroMaxTransac.Text)
vError:
End Sub

Private Sub txtGiroMaxTransac_LostFocus()
On Error GoTo vError
  txtGiroMaxTransac.Text = Format(CCur(txtGiroMaxTransac.Text), "Standard")
vError:
End Sub




Private Sub txtGiroMinimo_GotFocus()
On Error GoTo vError
  txtGiroMinimo.Text = CCur(txtGiroMinimo.Text)
vError:
End Sub

Private Sub txtGiroMinimo_LostFocus()
On Error GoTo vError
  txtGiroMinimo.Text = Format(CCur(txtGiroMinimo.Text), "Standard")
vError:
End Sub

Private Sub txtGirosMntTraslado_GotFocus()
On Error GoTo vError
  txtGirosMntTraslado.Text = CCur(txtGirosMntTraslado.Text)
vError:
End Sub

Private Sub txtGirosMntTraslado_LostFocus()
On Error GoTo vError
  txtGirosMntTraslado.Text = Format(CCur(txtGirosMntTraslado.Text), "Standard")
vError:
End Sub

Private Sub txtOficina_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select COD_OFICINA,Descripcion from SIF_OFICINAS"
        gBusquedas.Columna = "COD_OFICINA"
        gBusquedas.Orden = "COD_OFICINA"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtOficina.Text = gBusquedas.Resultado
           txtOficinaDesc.Text = gBusquedas.Resultado2
        End If
End If
End Sub


Private Sub txtOficinaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select COD_OFICINA,Descripcion from SIF_OFICINAS"
        gBusquedas.Columna = "COD_OFICINA"
        gBusquedas.Orden = "COD_OFICINA"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtOficina.Text = gBusquedas.Resultado
           txtOficinaDesc.Text = gBusquedas.Resultado2
        End If
End If
End Sub

Private Sub txtReserva_MontoMin_GotFocus()
On Error GoTo vError
  txtReserva_MontoMin.Text = CCur(txtReserva_MontoMin.Text)
vError:
End Sub

Private Sub txtReserva_MontoMin_LostFocus()
On Error GoTo vError
  txtReserva_MontoMin.Text = Format(CCur(txtReserva_MontoMin.Text), "Standard")
vError:
End Sub

Private Sub txtReserva_Plan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select cod_Plan,Descripcion from fnd_Planes"
        gBusquedas.Columna = "cod_Plan"
        gBusquedas.Orden = "cod_Plan"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtReserva_Plan.Text = gBusquedas.Resultado
           txtReserva_PlanDesc.Text = gBusquedas.Resultado2
        End If
End If
End Sub


Private Sub txtReserva_PlanDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select cod_Plan,Descripcion from fnd_Planes"
        gBusquedas.Columna = "Descripcion"
        gBusquedas.Orden = "Descripcion"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtReserva_Plan.Text = gBusquedas.Resultado
           txtReserva_PlanDesc.Text = gBusquedas.Resultado2
        End If
End If
End Sub

Private Sub txtRevPlanAhorro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select cod_Plan,Descripcion from fnd_Planes"
        gBusquedas.Columna = "cod_Plan"
        gBusquedas.Orden = "cod_Plan"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtRevPlanAhorro.Text = gBusquedas.Resultado
           txtRevPlanDesc.Text = gBusquedas.Resultado2
        End If
End If
End Sub

Private Sub txtRevPlanDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
        gBusquedas.Consulta = "select cod_Plan,Descripcion from fnd_Planes"
        gBusquedas.Columna = "Descripcion"
        gBusquedas.Orden = "Descripcion"
        gBusquedas.Filtro = ""
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        
        frmBusquedas.Show vbModal
        If gBusquedas.Resultado <> "" Then
           txtRevPlanAhorro.Text = gBusquedas.Resultado
           txtRevPlanDesc.Text = gBusquedas.Resultado2
        End If
End If

End Sub


Private Sub txtSupervisionMonto_GotFocus()
On Error GoTo vError
  txtSupervisionMonto.Text = CCur(txtSupervisionMonto.Text)
vError:
End Sub

Private Sub txtSupervisionMonto_LostFocus()
On Error GoTo vError
  txtSupervisionMonto.Text = Format(CCur(txtSupervisionMonto.Text), "Standard")
vError:
End Sub


Private Sub vgDocAdjunto_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vMovimiento As String
Dim vTempo As Integer, pCodigo As String

If vPaso Then Exit Sub

pCodigo = Trim(txtCodigoCorriente.Text)

With vgDocAdjunto

     .Row = Row
     .col = col
     
     If col = 4 Then 'Ultima Columna
        If .Value = 1 Then
           .col = 3
           vTempo = .Value
           .col = 1
           vMovimiento = "Registra"
           strSQL = "insert CRD_CATALOGO_ADJUNTOS(codigo,COD_ADJUNTO,opcional, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
                  & pCodigo & "','" & .Text & "'," & vTempo & ", '" & glogon.Usuario & "', dbo.mygetdate())"
        Else
           .col = 1
           vMovimiento = "Borrar"
           strSQL = "delete CRD_CATALOGO_ADJUNTOS where codigo = '" _
                  & pCodigo & "' and COD_ADJUNTO = '" & .Text & "'"
           
         End If
         
         Call ConectionExecute(strSQL)
         Call Bitacora(vMovimiento, "Catalogo, Adjunto: " & .Text & " a la Línea: " & pCodigo)
     End If
  
     If col = 3 Then 'Columna de Opcional
        .col = 3
        vTempo = .Value
        .col = 4
        If .Value = 1 Then
            .col = 1
            vMovimiento = "Modifica"
            strSQL = "update CRD_CATALOGO_ADJUNTOS set Opcional = " & vTempo & " where codigo = '" _
                   & pCodigo & "' and COD_ADJUNTO = '" & .Text & "'"
            
            Call ConectionExecute(strSQL)
            Call Bitacora(vMovimiento, "Catalogo, Adjunto: " & .Text & " a la Línea: " & pCodigo)
        End If
     End If
  
End With


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long

vGrid.Sheet = vGrid.ActiveSheet
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  Select Case vGrid.ActiveSheet
    Case 1 'Rangos
          i = fxGuardarRango
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
    
    Case 2 'Plazos / Tasa
        i = fxGuardarRangoPlazo
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
    
    Case 3 'Garantias
         With vGrid
           .Sheet = 3
           
           .Row = .ActiveRow
           
           
           .col = 2
           strSQL = "update crd_catalogo_garantias set utiliza_tasa_Garantia = " & .Value
           .col = 3
           strSQL = strSQL & ",Tasa_Garantia = " & IIf((.Text = ""), 0, .Text)
           .col = 4
           strSQL = strSQL & ",Utiliza_Tasa_Piso = " & .Value
           .col = 5
           strSQL = strSQL & ",Tasa_Piso = " & IIf((.Text = ""), 0, .Text)

           .col = 6
           strSQL = strSQL & ",utiliza_tasa_Techo = " & .Value
           .col = 7
           strSQL = strSQL & ",Tasa_Techo = " & IIf((.Text = ""), 0, .Text)
           
           .col = 8
           strSQL = strSQL & ",utiliza_maximos = " & .Value
           .col = 9
           strSQL = strSQL & ",Max_Monto = " & CCur(IIf((.Text = ""), 0, .Text)) _
        
           .col = 10 'Liquidez
           strSQL = strSQL & ",Liquidez_Minima = " & IIf((.Text = ""), 0, .Text)
        
           strSQL = strSQL & ",actualiza_fecha = dbo.MyGetdate(), actualiza_usuario ='" & glogon.Usuario & "'"
                    
           .col = 1
           strSQL = strSQL & " Where Garantia = '" & .CellTag & "' and codigo = '" & txtCodigoCorriente.Text & "'"
           
           Call ConectionExecute(strSQL)
        
           Call Bitacora("Modifica", "Garantia :" & .Text & " Linea : " & txtCodigoCorriente.Text)
         End With
    
  End Select

End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) _
    And (vGrid.ActiveSheet = 1 Or vGrid.ActiveSheet = 2) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If



End Sub



Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Call sbCargaRangos(NewSheet)
End Sub

Private Sub vGridAsg_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vMovimiento As String
Dim vTempo As Integer


If vPaso Then Exit Sub

On Error GoTo vError

With vGridAsg

Select Case .ActiveSheet
  Case 1 'Destinos
     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     If .Value = 1 Then
        .col = 1
        vMovimiento = "Registra"
        strSQL = "insert catalogo_DestinosAsg(codigo,cod_destino) values('" _
               & txtCodigoCorriente & "','" & .Text & "')"
     Else
        .col = 1
        vMovimiento = "Borrar"
        strSQL = "delete catalogo_DestinosAsg where codigo = '" _
               & txtCodigoCorriente & "' and cod_destino = '" & .Text & "'"
        
      End If
      Call ConectionExecute(strSQL)
      Call Bitacora(vMovimiento, "Destino : " & .Text & " a la Línea: " & txtCodigoCorriente)

  
  Case 2 'Cargos

     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     If .Value = 1 Then
       .col = 1
       vMovimiento = "Registra"
       strSQL = "insert cargos_asignacion(codigo,cod_cargo) values('" _
              & txtCodigoCorriente & "','" & .Text & "')"
     Else
       .col = 1
       vMovimiento = "Borrar"
       strSQL = "delete cargos_asignacion where codigo = '" _
              & txtCodigoCorriente & "' and cod_cargo = '" & .Text & "'"
       
     End If
     Call ConectionExecute(strSQL)
     Call Bitacora(vMovimiento, "Cargo : " & .Text & " a la Línea: " & txtCodigoCorriente)
          
  
  Case 3 'Requisitos
  
     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     
     If col = 4 Then 'Ultima Columna
        If .Value = 1 Then
           .col = 3
           vTempo = .Value
           .col = 1
           vMovimiento = "Registra"
           strSQL = "insert requisitos_asignacion(codigo,cod_requisito,opcional) values('" _
                  & txtCodigoCorriente & "','" & .Text & "'," & vTempo & ")"
        Else
           .col = 1
           vMovimiento = "Borrar"
           strSQL = "delete requisitos_asignacion where codigo = '" _
                  & txtCodigoCorriente & "' and cod_requisito = '" & .Text & "'"
           
         End If
         
         Call ConectionExecute(strSQL)
         Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Línea: " & txtCodigoCorriente)
     End If
  
     If col = 3 Then 'Columna de Opcional
        .col = 3
        vTempo = .Value
        .col = 4
        If .Value = 1 Then
            .col = 1
            vMovimiento = "Modifica"
            strSQL = "update requisitos_asignacion set Opcional = " & vTempo & " where codigo = '" _
                   & txtCodigoCorriente & "' and cod_requisito = '" & .Text & "'"
            
            Call ConectionExecute(strSQL)
            Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Línea: " & txtCodigoCorriente)
        End If
     End If
  
  
  Case 4 'Recursos
     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     If .Value = 1 Then
        .col = 1
        vMovimiento = "Registra"
        strSQL = "insert catalogo_asignaGrp(codigo,cod_grupo) values('" _
               & txtCodigoCorriente & "','" & .Text & "')"
     Else
        .col = 1
        vMovimiento = "Borrar"
        strSQL = "delete catalogo_asignaGrp where codigo = '" _
               & txtCodigoCorriente & "' and cod_grupo = '" & .Text & "'"
        
      End If
      Call ConectionExecute(strSQL)
      Call Bitacora(vMovimiento, "Recurso : " & .Text & " a la Línea: " & txtCodigoCorriente)


  Case 5 'Cartera
     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     If .Value = 1 Then
        .col = 1
        vMovimiento = "Registra"
        strSQL = "insert CBR_CLASIFICACION_DETALLE(codigo,COD_CLASIFICACION) values('" _
               & txtCodigoCorriente & "','" & .Text & "')"
     Else
        .col = 1
        vMovimiento = "Borrar"
        strSQL = "delete CBR_CLASIFICACION_DETALLE where codigo = '" _
               & txtCodigoCorriente & "' and COD_CLASIFICACION = '" & .Text & "'"
        
      End If
      Call ConectionExecute(strSQL)
      Call Bitacora(vMovimiento, "Cartera de Mora : " & .Text & " a la Línea: " & txtCodigoCorriente)


  Case 6 'Refundibles
     .Sheet = .ActiveSheet
     .Row = Row
     .col = col
     
     If .Value = 1 Then
        .col = 1
        vMovimiento = "Registra"
        strSQL = "insert CRD_CATALOGO_REFUNDIBLES(codigo,COD_REFUNDIBLE, REGISTRO_FECHA, REGISTRO_USUARIO) values('" _
               & txtCodigoCorriente & "','" & .Text & "', dbo.mygetdate(),'" & glogon.Usuario & "')"
     Else
        .col = 1
        vMovimiento = "Borrar"
        strSQL = "delete CRD_CATALOGO_REFUNDIBLES where codigo = '" _
               & txtCodigoCorriente & "' and COD_REFUNDIBLE = '" & .Text & "'"
        
      End If
      Call ConectionExecute(strSQL)
      Call Bitacora(vMovimiento, "Lista Refundibles: " & .Text & " a la Línea: " & txtCodigoCorriente)
End Select

End With


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub vGridAsg_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
    Call sbAsignaciones(NewSheet)
End Sub
