VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Begin VB.Form frmFSL_Expediente 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente FOSOL"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7215
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   12975
      _Version        =   1441793
      _ExtentX        =   22886
      _ExtentY        =   12726
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
      Item(0).Caption =   "General"
      Item(0).ControlCount=   28
      Item(0).Control(0)=   "scTitulos(0)"
      Item(0).Control(1)=   "scTitulos(1)"
      Item(0).Control(2)=   "label5(0)"
      Item(0).Control(3)=   "txtPresentaCedula"
      Item(0).Control(4)=   "txtPresentaNombre"
      Item(0).Control(5)=   "label5(1)"
      Item(0).Control(6)=   "txtPresentaNotas"
      Item(0).Control(7)=   "label5(2)"
      Item(0).Control(8)=   "cboComite"
      Item(0).Control(9)=   "dtpRefFecha"
      Item(0).Control(10)=   "txtRefNumero"
      Item(0).Control(11)=   "cboRefTipoDoc"
      Item(0).Control(12)=   "cboTipo"
      Item(0).Control(13)=   "cboCausa"
      Item(0).Control(14)=   "cboEnfermedad"
      Item(0).Control(15)=   "dtpEnfermedad"
      Item(0).Control(16)=   "label5(3)"
      Item(0).Control(17)=   "label5(4)"
      Item(0).Control(18)=   "label5(5)"
      Item(0).Control(19)=   "label5(6)"
      Item(0).Control(20)=   "label5(7)"
      Item(0).Control(21)=   "label5(8)"
      Item(0).Control(22)=   "label5(9)"
      Item(0).Control(23)=   "label5(10)"
      Item(0).Control(24)=   "label5(11)"
      Item(0).Control(25)=   "label5(12)"
      Item(0).Control(26)=   "txtNotas"
      Item(0).Control(27)=   "txtEnfermedadNotas"
      Item(1).Caption =   "Requisitos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lswRequisitos"
      Item(1).Control(1)=   "Label3(14)"
      Item(2).Caption =   "Operaciones"
      Item(2).ControlCount=   13
      Item(2).Control(0)=   "txtTD_Texto_01"
      Item(2).Control(1)=   "txtTD_Texto_02"
      Item(2).Control(2)=   "txtTD_Destino"
      Item(2).Control(3)=   "txtTotalDisponible"
      Item(2).Control(4)=   "txtLiquidacionMonto"
      Item(2).Control(5)=   "txtTotalAplicado"
      Item(2).Control(6)=   "vgCreditos"
      Item(2).Control(7)=   "lblTD_Label_01"
      Item(2).Control(8)=   "lblTD_Label_02"
      Item(2).Control(9)=   "Label7(2)"
      Item(2).Control(10)=   "Label7(4)"
      Item(2).Control(11)=   "Label7(9)"
      Item(2).Control(12)=   "Label7(12)"
      Item(3).Caption =   "Resolución"
      Item(3).ControlCount=   14
      Item(3).Control(0)=   "fraValidaMiembro"
      Item(3).Control(1)=   "cboResolucion"
      Item(3).Control(2)=   "txtResolucionNotas"
      Item(3).Control(3)=   "tlbResolucion"
      Item(3).Control(4)=   "imgExpedientesActivos"
      Item(3).Control(5)=   "imgTiempoPresentacion"
      Item(3).Control(6)=   "imgRequisitos"
      Item(3).Control(7)=   "Label3(5)"
      Item(3).Control(8)=   "Label3(4)"
      Item(3).Control(9)=   "Label3(3)"
      Item(3).Control(10)=   "Label3(13)"
      Item(3).Control(11)=   "Label3(12)"
      Item(3).Control(12)=   "Label3(11)"
      Item(3).Control(13)=   "lswComite"
      Item(4).Caption =   "Gestiones"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "vgGestiones"
      Item(5).Caption =   "Apelaciones"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "vgApelaciones"
      Begin XtremeSuiteControls.ListView lswRequisitos 
         Height          =   6495
         Left            =   -66040
         TabIndex        =   72
         Top             =   600
         Visible         =   0   'False
         Width           =   8895
         _Version        =   1441793
         _ExtentX        =   15690
         _ExtentY        =   11456
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
      Begin XtremeSuiteControls.ListView lswComite 
         Height          =   4695
         Left            =   -64480
         TabIndex        =   74
         Top             =   960
         Visible         =   0   'False
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   8281
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
      Begin VB.Frame fraValidaMiembro 
         Caption         =   "Validación del Miembro de Comité"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -66520
         TabIndex        =   51
         Top             =   1440
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtMiembroClave 
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
            IMEMode         =   3  'DISABLE
            Left            =   3480
            PasswordChar    =   "*"
            TabIndex        =   53
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtMiembroUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3480
            TabIndex        =   52
            ToolTipText     =   "Número de Tramite"
            Top             =   720
            Width           =   2175
         End
         Begin MSComctlLib.Toolbar tlbValidaMiembro 
            Height          =   330
            Left            =   3000
            TabIndex        =   54
            Top             =   1680
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   582
            ButtonWidth     =   1826
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgLista"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Valida"
                  Key             =   "Valida"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Cerrar"
                  Key             =   "Cerrar"
                  ImageIndex      =   3
               EndProperty
            EndProperty
            MousePointer    =   1
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario .: "
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
            Left            =   2040
            TabIndex        =   57
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Clave .: "
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
            Left            =   2040
            TabIndex        =   56
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   240
            X2              =   5760
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label lblMiembro 
            Alignment       =   1  'Right Justify
            Caption         =   "Miembro.: "
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
            Left            =   360
            TabIndex        =   55
            Top             =   360
            Width           =   5295
         End
      End
      Begin VB.ComboBox cboResolucion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -62440
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   6120
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox cboEnfermedad 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4680
         Width           =   6495
      End
      Begin VB.ComboBox cboCausa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4080
         Width           =   6495
      End
      Begin VB.ComboBox cboTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3720
         Width           =   6495
      End
      Begin VB.ComboBox cboRefTipoDoc 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ComboBox cboComite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3360
         Width           =   6495
      End
      Begin XtremeSuiteControls.FlatEdit txtPresentaNombre 
         Height          =   330
         Left            =   6240
         TabIndex        =   19
         Top             =   1080
         Width           =   6615
         _Version        =   1441793
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
      Begin FPSpreadADO.fpSpread vgApelaciones 
         Height          =   6615
         Left            =   -70000
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   12975
         _Version        =   524288
         _ExtentX        =   22886
         _ExtentY        =   11668
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
         MaxCols         =   7
         SpreadDesigner  =   "frmFSL_Expediente.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vgGestiones 
         Height          =   6615
         Left            =   -70000
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   12855
         _Version        =   524288
         _ExtentX        =   22675
         _ExtentY        =   11668
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
         MaxCols         =   4
         SpreadDesigner  =   "frmFSL_Expediente.frx":072F
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPresentaCedula 
         Height          =   330
         Left            =   1920
         TabIndex        =   18
         Top             =   1080
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPresentaNotas 
         Height          =   1050
         Left            =   1920
         TabIndex        =   21
         Top             =   1440
         Width           =   10935
         _Version        =   1441793
         _ExtentX        =   19288
         _ExtentY        =   1852
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1410
         Left            =   120
         TabIndex        =   38
         Top             =   5640
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
         _ExtentY        =   2487
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
      Begin XtremeSuiteControls.FlatEdit txtEnfermedadNotas 
         Height          =   1410
         Left            =   6480
         TabIndex        =   39
         Top             =   5640
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
         _ExtentY        =   2487
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
      Begin XtremeSuiteControls.DateTimePicker dtpEnfermedad 
         Height          =   330
         Left            =   1920
         TabIndex        =   40
         Top             =   4680
         Width           =   1935
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpRefFecha 
         Height          =   330
         Left            =   1920
         TabIndex        =   41
         Top             =   4080
         Width           =   1935
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtRefNumero 
         Height          =   330
         Left            =   1920
         TabIndex        =   42
         Top             =   3720
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   5415
         Left            =   -70000
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   12855
         _Version        =   524288
         _ExtentX        =   22675
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
         MaxCols         =   13
         SpreadDesigner  =   "frmFSL_Expediente.frx":0D81
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbResolucion 
         Height          =   330
         Left            =   -59200
         TabIndex        =   59
         Top             =   6120
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         ButtonWidth     =   2090
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgLista"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Guardar"
               Key             =   "Resolucion"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalDisponible 
         Height          =   330
         Left            =   -67480
         TabIndex        =   66
         Top             =   6000
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtTotalAplicado 
         Height          =   330
         Left            =   -67480
         TabIndex        =   67
         Top             =   6360
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtLiquidacionMonto 
         Height          =   330
         Left            =   -67480
         TabIndex        =   68
         Top             =   6720
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtTD_Destino 
         Height          =   330
         Left            =   -61480
         TabIndex        =   69
         Top             =   6000
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtTD_Texto_01 
         Height          =   330
         Left            =   -61480
         TabIndex        =   70
         Top             =   6360
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtTD_Texto_02 
         Height          =   330
         Left            =   -61480
         TabIndex        =   71
         Top             =   6720
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtResolucionNotas 
         Height          =   2610
         Left            =   -69880
         TabIndex        =   73
         Top             =   960
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   4604
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
      Begin VB.Label Label3 
         Caption         =   "Notas de la Resolución.:"
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
         Index           =   11
         Left            =   -69880
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Miembros del Comité.:"
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
         Index           =   12
         Left            =   -64480
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   " Resolución.:"
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
         Index           =   13
         Left            =   -64240
         TabIndex        =   63
         Top             =   6120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple Requisitos .:"
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
         Left            =   -69880
         TabIndex        =   62
         Top             =   3840
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple Tiempo de Presentación.:"
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
         Left            =   -69880
         TabIndex        =   61
         Top             =   4200
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cumple No. Expedientes Activos.:"
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
         Left            =   -69880
         TabIndex        =   60
         Top             =   4560
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Image imgRequisitos 
         Height          =   255
         Left            =   -66400
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgTiempoPresentacion 
         Height          =   255
         Left            =   -66400
         Top             =   4200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgExpedientesActivos 
         Height          =   255
         Left            =   -66400
         Top             =   4560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Total Aplicado"
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
         Height          =   255
         Index           =   12
         Left            =   -69280
         TabIndex        =   50
         Top             =   6360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "(Liquidación)"
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
         Height          =   255
         Index           =   9
         Left            =   -69280
         TabIndex        =   49
         Top             =   6720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Disponible"
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
         Height          =   255
         Index           =   4
         Left            =   -69280
         TabIndex        =   48
         Top             =   6000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Desembolso"
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
         Height          =   255
         Index           =   2
         Left            =   -63760
         TabIndex        =   47
         Top             =   6000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblTD_Label_02 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Height          =   255
         Left            =   -63520
         TabIndex        =   46
         Top             =   6720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblTD_Label_01 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
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
         Height          =   255
         Left            =   -63640
         TabIndex        =   45
         Top             =   6360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Cumplimiento con Requisitos.:"
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
         Index           =   14
         Left            =   -68920
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   12
         Left            =   6480
         TabIndex        =   37
         Top             =   5280
         Width           =   3975
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas de la enfermedad/causa.:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   36
         Top             =   5280
         Width           =   3975
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas del Expediente.:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   10
         Left            =   4560
         TabIndex        =   35
         Top             =   4680
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Enfermedad ...:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   34
         Top             =   4080
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Causa..:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   8
         Left            =   4560
         TabIndex        =   33
         Top             =   3720
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Plan...:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   7
         Left            =   4560
         TabIndex        =   32
         Top             =   3360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comité...:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   31
         Top             =   4560
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Fecha Diagnóstico Enfermedad"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   30
         Top             =   4080
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ref. Fecha.:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   29
         Top             =   3720
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ref. No.:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   28
         Top             =   3360
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Referencia...:"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Notas de la Presentación"
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   20
         Top             =   1080
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cédula"
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
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   16
         Top             =   2640
         Width           =   12975
         _Version        =   1441793
         _ExtentX        =   22886
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Datos del Expediente:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   4210752
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulos 
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   12975
         _Version        =   1441793
         _ExtentX        =   22886
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Datos del Solicitante:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   4210752
      End
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   13185
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   2955
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   3150
         TabIndex        =   2
         Top             =   30
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   582
         ButtonWidth     =   2937
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gestiones"
               Key             =   "Gestiones"
               Object.ToolTipText     =   "Registro de Gestiones Realizadas"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Apelación"
               Key             =   "Apelacion"
               Object.ToolTipText     =   "Apelación de la Resolución"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar (FOSOL)"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicación de la Resolución"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Boleta de Registro"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   480
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
            Picture         =   "frmFSL_Expediente.frx":17B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":18CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":19EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1D17
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":1F5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":2085
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":21AF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList imgLista 
      Left            =   11520
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":22A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":8B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Expediente.frx":8C03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   9135
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7832
            MinWidth        =   7832
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
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
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   600
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   873
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
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
      Left            =   4440
      TabIndex        =   11
      Top             =   1320
      Width           =   6975
      _Version        =   1441793
      _ExtentX        =   12303
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   330
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Cédula"
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Expediente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgEstado 
      Height          =   255
      Left            =   5160
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "frmFSL_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

fxValida = True
vMensaje = ""

If Not vEdita Then
    strSQL = "select dbo.fxFSL_ExpedienteValidaRegistro('" & txtCedula.Text & "','" & SIFGlobal.fxCodText(cboTipo.Text) _
           & "','" & SIFGlobal.fxCodText(cboCausa.Text) & "',0) as 'Cumple'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Cumple = 0 Then
        vMensaje = vMensaje & vbCrLf & "- El caso ya fue presentado anteriormente...verifique!"
    End If
    
End If
    
If Len(txtNotas.Text) <= 10 Then vMensaje = vMensaje & vbCrLf & "- Indique una Nota válida!"


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEnfermedad.SetFocus

End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus

End Sub

Private Sub cboEnfermedad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEnfermedadNotas.SetFocus

End Sub

Private Sub cboRefTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRefNumero.SetFocus

End Sub

Private Sub cboTipo_Click()
Dim strSQL As String

cboCausa.Clear

If vPaso Then Exit Sub
If cboTipo.ListCount = 0 Then Exit Sub

strSQL = "select rtrim(COD_CAUSA) + ' - ' + descripcion as 'ItmX'" _
       & " from FSL_PLANES_CAUSAS" _
       & " where COD_PLAN = '" & SIFGlobal.fxCodText(cboTipo.Text) _
       & "' order by COD_CAUSA"
Call sbLlenaCbo(cboCausa, strSQL, False, False)

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus
End Sub



Private Sub dtpRefFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboComite.SetFocus

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True


If txtExpediente.Text = "" Then txtExpediente.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 cod_Expediente from FSL_Expedientes"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where cod_Expediente > " & txtExpediente & " order by cod_Expediente asc"
Else
   strSQL = strSQL & " where cod_Expediente < " & txtExpediente & " order by cod_Expediente desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtExpediente.Text = rs!COD_EXPEDIENTE
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 7
 
With lswRequisitos.ColumnHeaders
    .Clear
    .Add , , "Requisito", 6500
    .Add , , "Opcional", lswRequisitos.Width - 6800
End With
 
 
With lswComite.ColumnHeaders
    .Clear
    .Add , , "Nombre", 4000
    .Add , , "", lswComite.Width - 4300
End With

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "activo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 

Call sbPantallaInicial
Call sbLimpiaPantalla


End Sub

Private Sub sbPantallaInicial()
Dim strSQL As String

Me.MousePointer = vbHourglass
vPaso = True

strSQL = "select Rtrim(COD_PLAN) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_PLANES where activo = 1"
Call sbLlenaCbo(cboTipo, strSQL, False)

strSQL = "select Rtrim(COD_COMITE) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_COMITES where activo = 1"
Call sbLlenaCbo(cboComite, strSQL, False)

strSQL = "select Rtrim(COD_ENFERMEDAD) + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  FSL_TIPOS_ENFERMEDADES where activa = 1"
Call sbLlenaCbo(cboEnfermedad, strSQL, False)

cboRefTipoDoc.Clear
cboRefTipoDoc.AddItem "Accion de personal"
cboRefTipoDoc.AddItem "Acta defunción"
cboRefTipoDoc.Text = "Accion de personal"

cboResolucion.Clear
cboResolucion.AddItem "Aprobado"
cboResolucion.AddItem "Rechazado"
cboResolucion.Text = "Aprobado"


vPaso = False
Call cboTipo_Click

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

Me.MousePointer = vbHourglass
vPaso = True
 tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
 tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
 tlbAux.Buttons.Item(4).Enabled = False 'Aplicar

   
tcMain.Item(1).Enabled = False
tcMain.Item(2).Enabled = False
tcMain.Item(3).Enabled = False
tcMain.Item(4).Enabled = False
tcMain.Item(5).Enabled = False
   
 tcMain.Item(0).Selected = True

Set imgEstado.Picture = ImageList1.ListImages.Item(6).Picture
Set imgRequisitos.Picture = ImageList1.ListImages.Item(6).Picture
Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(6).Picture
Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(6).Picture

txtExpediente.Text = ""
txtEstado.Text = "Pendiente"

txtCedula.Text = ""
txtNombre.Text = ""

txtPresentaCedula.Text = ""
txtPresentaNombre.Text = ""
txtPresentaNotas.Text = ""

txtNotas.Text = ""
txtEnfermedadNotas.Text = ""

 dtpRefFecha.Value = fxFechaServidor
 dtpEnfermedad.Value = dtpRefFecha.Value
 txtRefNumero.Text = ""
 
 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 
vPaso = False

Me.MousePointer = vbDefault

End Sub



Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

strSQL = "select Ex.*, Soc.NOMBRE" _
        & ", Pl.COD_PLAN + ' - ' + rtrim(Pl.DESCRIPCION) as 'Plan'" _
        & ", Pc.COD_CAUSA + ' - ' + rtrim(Pc.DESCRIPCION) as 'Causa'" _
        & ", Te.COD_ENFERMEDAD + ' - ' + rtrim(Te.DESCRIPCION) as 'Enfermedad'" _
        & ", Co.COD_COMITE + ' - ' + rtrim(Co.DESCRIPCION) as 'Comite'" _
        & " from FSL_EXPEDIENTES Ex" _
        & " inner join SOCIOS Soc on Ex.Cedula = Soc.Cedula" _
        & " inner join FSL_PLANES Pl on Ex.COD_PLAN = Pl.COD_PLAN" _
        & " inner join FSL_PLANES_CAUSAS Pc on Ex.COD_PLAN = Pc.COD_PLAN and Ex.COD_CAUSA = Pc.COD_CAUSA" _
        & " inner join FSL_TIPOS_ENFERMEDADES Te on Ex.COD_ENFERMEDAD = Te.COD_ENFERMEDAD" _
        & " inner join FSL_COMITES Co on Ex.COD_COMITE = Co.COD_COMITE" _
        & " Where Ex.COD_EXPEDIENTE = " & txtExpediente.Text

Call OpenRecordSet(rs, strSQL)

Call sbLimpiaPantalla

If Not rs.EOF And Not rs.BOF Then

  Call sbToolBar(tlb, "activo")
  vEdita = True

  txtExpediente.Text = rs!COD_EXPEDIENTE
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre

  
  txtPresentaCedula.Text = rs!Presenta_Cedula & ""
  txtPresentaNombre.Text = rs!PRESENTA_NOMBRE & ""
  txtPresentaNotas.Text = rs!Presenta_Notas & ""
  
  txtRefNumero.Text = rs!Referencia_Numero
  dtpRefFecha.Value = rs!Fecha_Establece_Causa
  
  dtpEnfermedad.Value = rs!Enfermedad_Fecha
  
  
  txtNotas.Text = rs!notas & ""
  txtEnfermedadNotas.Text = rs!Enfermedad_Notas & ""
       
       
  vPaso = True
        Call sbCboAsignaDato(cboTipo, rs!Plan, True)
        Call sbCboAsignaDato(cboCausa, rs!Causa, True)
        Call sbCboAsignaDato(cboComite, rs!Comite, True)
        Call sbCboAsignaDato(cboEnfermedad, rs!Enfermedad, True)
        Call sbCboAsignaDato(cboRefTipoDoc, rs!Referencia_Documento, True)
  vPaso = False
  

    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    tcMain.Item(3).Enabled = True
    tcMain.Item(4).Enabled = True
    tcMain.Item(5).Enabled = True


  tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
  tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
  tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
  
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
        txtEstado.Tag = "P"
        Set imgEstado.Picture = ImageList1.ListImages.Item(8).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
     
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
        txtEstado.Tag = "A"
        Set imgEstado.Picture = ImageList1.ListImages.Item(7).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = True 'Aplicar
   
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
        txtEstado.Tag = "R"
        Set imgEstado.Picture = ImageList1.ListImages.Item(9).Picture
        
        tlbAux.Buttons.Item(1).Enabled = True 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = True 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
   
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
        txtEstado.Tag = "X"
        Set imgEstado.Picture = ImageList1.ListImages.Item(10).Picture
        
        tlbAux.Buttons.Item(1).Enabled = False 'Gestiones
        tlbAux.Buttons.Item(2).Enabled = False 'Apelacion
        tlbAux.Buttons.Item(4).Enabled = False 'Aplicar
 
  End Select


 'Totales (Operaciones)
 If rs!TIPO_DESEMBOLSO = "F" Then
    lblTD_Label_01.Caption = "Plan"
    lblTD_Label_02.Caption = "Contrato"
    txtTD_Destino.Text = "-- FONDOS --"
    txtTD_Texto_01.Text = rs!FND_PLAN & ""
    txtTD_Texto_02.Text = rs!FND_CONTRATO & ""
 Else
    lblTD_Label_01.Caption = "No. Solicitud"
    lblTD_Label_02.Caption = "No. Remesa"
    txtTD_Destino.Text = "-- TESORERIA --"
    txtTD_Texto_01.Text = rs!Tesoreria_Solicitud & ""
    txtTD_Texto_02.Text = rs!TESORERIA_REMESA & ""
 End If
  
 txtTotalAplicado.Text = Format(rs!Total_Aplicado, "Standard")
 txtTotalDisponible.Text = Format(rs!Total_Disponible, "Standard")
 txtLiquidacionMonto.Text = Format(rs!TOTAL_SOBRANTE, "Standard")
  
 txtTotalAplicado.ToolTipText = rs!Apl_Tipo_Doc & " .: " & rs!Apl_Num_Doc
  
  
 'Fin de Totales
  txtResolucionNotas.Text = rs!Resolucion_Notas & ""

  Select Case rs!Resolucion_Estado
    Case "A" 'Aprobado
        cboResolucion.Text = "Aprobado"
    Case "R" 'Rechazado
        cboResolucion.Text = "Rechazado"
  End Select
    
    StatusBarX.Panels(1).Text = "Resolución.:" & rs!Resolucion_Fecha & " ¦ " & rs!Resolucion_Usuario
    StatusBarX.Panels(2).Text = "Rf.:" & rs!registro_Fecha & ""
    StatusBarX.Panels(3).Text = "Ru.: " & rs!Registro_Usuario & ""

Else
    MsgBox "No existe el expediente, verifique!", vbCritical
End If
rs.Close

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub lswComite_DblClick()
Dim strSQL As String, rs As New ADODB.Recordset

If lswComite.ListItems.Count = 0 Then Exit Sub


With lswComite.SelectedItem
  If .SubItems(1) = "Pendiente" Then
     fraValidaMiembro.Visible = True
     lblMiembro.Caption = .Text
     lblMiembro.Tag = .Tag

     strSQL = "select usuario_Vinculado from FSL_Comites_Miembros" _
           & " where cedula = '" & .Tag & "' and cod_comite = '" & SIFGlobal.fxCodText(cboComite.Text) & "'"
     Call OpenRecordSet(rs, strSQL)
         txtMiembroUsuario.Text = rs!Usuario_Vinculado
     rs.Close
     
     txtMiembroClave.SetFocus

  End If
  
End With

End Sub




Private Sub sbReporte()

 With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes FOSOL"
 
  .Connect = glogon.ConectRPT
   
    
   
  .ReportFileName = SIFGlobal.fxPathReportes("FSL_ExpedienteBoleta.rpt")
  
  .Formulas(0) = "fxCodigoBarras = '*" & txtExpediente.Text & "*'"
  .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(2) = "fxUsuario='USUARIO: " & glogon.Usuario & "'"
  .Formulas(3) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                   
  .SelectionFormula = "{vFSL_CasosLista.COD_EXPEDIENTE} =" & txtExpediente.Text

  

  .PrintReport
  
  
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswRequisitos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If txtEstado.Tag <> "P" Then
    MsgBox "El expediente no está Pendiente! No se pueden modificar los requisitos!", vbExclamation
    Exit Sub
End If

strSQL = "update FSL_EXPEDIENTES_REQUISITOS set Estado = " & IIf(Item.Checked, 1, 0) _
       & ", registro_fecha = getdate(), registro_usuario = '" & glogon.Usuario & "'" _
       & " where cod_expediente = " & txtExpediente.Text & " and Cod_Requisito = '" & Item.Tag & "'"

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If Not IsNumeric(txtExpediente.Text) Then
    tcMain.Item(0).Selected = True
End If

Select Case Item.Index   ' ssTab.Tab
  Case 0 'General
  Case 1 'Requisitos
       vPaso = True
       
       strSQL = "Select Ex.COD_REQUISITO,Rq.DESCRIPCION, EX.Estado, Ex.Opcional " _
              & " from FSL_EXPEDIENTES_REQUISITOS Ex " _
              & "  inner join FSL_REQUISITOS Rq on Ex.cod_requisito = Rq.cod_requisito" _
              & " where Ex.cod_Expediente = " & txtExpediente.Text
        Call OpenRecordSet(rs, strSQL)
        With lswRequisitos.ListItems
           .Clear
           Do While Not rs.EOF
            Set itmX = .Add(, , rs!Descripcion)
                itmX.Tag = rs!Cod_Requisito
                If rs!Opcional = 1 Then
                    itmX.SubItems(1) = "Sí"
                Else
                    itmX.SubItems(1) = "No"
                End If
                itmX.Checked = rs!Estado
            rs.MoveNext
           Loop
        End With
        rs.Close
        
        vPaso = False
        
  Case 2 'Operaciones
        strSQL = "select E.ID_SOLICITUD, E.REFERENCIA, Gar.DESCRIPCION, R.PRIDEDUC, R.MONTOAPR " _
               & "   , E.SALDO_CORTE , E.MONTO_BASE, E.PORC_RELACION , E.TIPO_TABLA , E.PORCENTAJE" _
               & "   , E.MONTO_RECONOCIMIENTO, E.TIEMPO_TRANS, case when E.Tipo_Base = 'S' then 'Saldo' else 'Mnt.Form.' end as 'BASE'" _
               & " from FSL_EXPEDIENTES_DETALLE  E inner join REG_CREDITOS R on E.ID_SOLICITUD = R.ID_SOLICITUD" _
               & " inner join CRD_GARANTIA_TIPOS Gar on R.GARANTIA = Gar.GARANTIA" _
               & " Where E.COD_EXPEDIENTE = " & txtExpediente.Text & " Order by isnull(E.referencia,E.id_Solicitud) desc"
        Call sbCargaGrid(vgCreditos, 13, strSQL, True)
  
  
  Case 3 'Resolucion
       strSQL = "Select Cm.Cedula,Cm.Nombre,isnull(Ec.Asigna_Usuario,'No!') as 'ASIGNADO'" _
              & " from FSL_EXPEDIENTES Ex" _
              & "  inner join FSL_COMITES_MIEMBROS Cm on Ex.COD_COMITE = Cm.COD_COMITE" _
              & "   left join FSL_EXPEDIENTE_COMITE Ec on Ex.COD_EXPEDIENTE = Ec.COD_EXPEDIENTE and Ex.COD_COMITE = Ec.COD_COMITE" _
              & "         and Cm.Cedula = Ex.Cedula" _
              & " where Ex.cod_Expediente = " & txtExpediente.Text & " and Cm.Activo = 1"
        Call OpenRecordSet(rs, strSQL)
        
        vPaso = True
        
        With lswComite.ListItems
           .Clear
           Do While Not rs.EOF
            Set itmX = .Add(, , rs!Nombre)
                itmX.Tag = rs!Cedula
                If rs!Asignado <> "No!" Then
                   itmX.Checked = True
                   itmX.SubItems(1) = "Validado"
                Else
                   itmX.SubItems(1) = "Pendiente"
                End If
            rs.MoveNext
           Loop
        End With
        rs.Close
  
        vPaso = False
  
        'Refresca Validaciones

        strSQL = "select dbo.fxFSL_ExpedienteValidaRequisitos(Ex.Cod_Expediente) as 'CumpleRequisitos'" _
                & ", dbo.fxFSL_ExpedienteValidaTiempoPresentacion(Ex.Cod_Expediente) as 'CumpleTiempo'" _
                & ", dbo.fxFSL_ExpedienteValidaRegistro(Ex.Cedula, Ex.Cod_Plan, Ex.Cod_Causa,Ex.Cod_Expediente) as 'CumpleRegistro'" _
                & " from FSL_EXPEDIENTES Ex" _
                & " Where Ex.COD_EXPEDIENTE = " & txtExpediente.Text
        Call OpenRecordSet(rs, strSQL)
            
            imgRequisitos.Tag = rs!CumpleRequisitos
            imgTiempoPresentacion.Tag = rs!CumpleTiempo
            imgExpedientesActivos.Tag = rs!CumpleRegistro
            
            If rs!CumpleRequisitos = 1 Then
               Set imgRequisitos.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgRequisitos.Picture = ImageList1.ListImages.Item(9).Picture
            End If
            
            
            If rs!CumpleTiempo = 1 Then
               Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgTiempoPresentacion.Picture = ImageList1.ListImages.Item(9).Picture
            End If
            
            If rs!CumpleRegistro = 1 Then
               Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(7).Picture
            Else
               Set imgExpedientesActivos.Picture = ImageList1.ListImages.Item(9).Picture
            End If
   
        rs.Close
  
  Case 4 'Gestiones
        strSQL = "select Tg.Descripcion, Eg.*" _
               & " from FSL_EXPEDIENTE_GESTIONES Eg inner join FSL_TIPOS_GESTIONES Tg on Eg.COD_GESTION = Tg.COD_GESTION" _
               & " Where Eg.cod_Expediente = " & txtExpediente.Text & " order by registro_fecha desc"
        Call OpenRecordSet(rs, strSQL)
        vgGestiones.MaxRows = 0
        Do While Not rs.EOF
          vgGestiones.MaxRows = vgGestiones.MaxRows + 1
          vgGestiones.Row = vgGestiones.MaxRows
          
          vgGestiones.Col = 1
          vgGestiones.Text = rs!Descripcion
          vgGestiones.TextTip = TextTipFixed
          vgGestiones.TextTipDelay = 1000
        
          vgGestiones.CellNote = "Fecha : " & rs!registro_Fecha & vbCrLf & "Usuario : " & rs!Registro_Usuario
          vgGestiones.CellTag = CStr(rs!Linea)
            
          vgGestiones.Col = 2
          vgGestiones.Text = rs!notas
              
          vgGestiones.Col = 3
          vgGestiones.Text = rs!registro_Fecha
              
          vgGestiones.Col = 4
          vgGestiones.Text = rs!Registro_Usuario
         
          vgGestiones.RowHeight(vgGestiones.Row) = vgGestiones.MaxTextRowHeight(vgGestiones.Row)
          
         rs.MoveNext
        Loop
        rs.Close
  
  
  Case 5 'Apelaciones

        strSQL = "select Ta.Descripcion, Ea.*" _
               & " from FSL_EXPEDIENTES_APELACIONES Ea inner join FSL_TIPOS_APELACIONES Ta on Ea.COD_APELACION = Ta.COD_APELACION" _
               & " Where Ea.cod_Expediente = " & txtExpediente.Text & " order by registra_fecha desc"
        
        Call OpenRecordSet(rs, strSQL)
        vgApelaciones.MaxRows = 0
        
        Do While Not rs.EOF
          vgApelaciones.MaxRows = vgApelaciones.MaxRows + 1
          vgApelaciones.Row = vgApelaciones.MaxRows
          
          vgApelaciones.Col = 1
          vgApelaciones.Text = rs!Descripcion
        
          vgApelaciones.CellNote = "Fecha : " & rs!Registra_Fecha & vbCrLf & "Usuario : " & rs!Registra_Usuario
          vgApelaciones.CellTag = CStr(rs!Linea)
           
            
          vgApelaciones.Col = 2
          vgApelaciones.Text = rs!notas
              
          vgApelaciones.Col = 3
          
          Select Case rs!Resolucion
             Case "P" 'Pendiente
                  vgApelaciones.Text = "Pendiente"
             Case "A" 'Aprobada
                  vgApelaciones.Text = "Aprobada"
             Case "R" 'Rechazada
                  vgApelaciones.Text = "Rechazada"
          End Select
          
          vgApelaciones.Col = 4
          vgApelaciones.Text = rs!Registra_Fecha
              
          vgApelaciones.Col = 5
          vgApelaciones.Text = rs!Registra_Usuario
              
              
          vgApelaciones.Col = 6
          vgApelaciones.Text = rs!Presenta_Identificacion & ""
              
          vgApelaciones.Col = 7
          vgApelaciones.Text = rs!PRESENTA_NOMBRE & ""
              
              
          vgApelaciones.RowHeight(vgApelaciones.Row) = vgApelaciones.MaxTextRowHeight(vgApelaciones.Row)
          
         rs.MoveNext
        Loop
        rs.Close

End Select


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtPresentaCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtExpediente.Text = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "Ex.Cedula"
       gBusquedas.Orden = "Ex.Cedula"
       gBusquedas.Consulta = "select Ex.cod_expediente,Soc.nombre from FSL_EXPEDIENTES Ex inner join SOCIOS Soc on EX.cedula = Soc.Cedula"
       frmBusquedas.Show vbModal
       txtExpediente.SetFocus
       txtExpediente = gBusquedas.Resultado

    Case "REPORTES"
       Call sbReporte
       
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub tlbResolucion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim y As Integer, vNumResolutores As Integer

On Error GoTo vError

If txtEstado.Tag <> "P" Then
   MsgBox "Este Expediente No se encuentra pendiente, no puede registrarse una resolución!", vbExclamation
   Exit Sub
End If

If Mid(cboResolucion.Text, 1, 1) = "A" Then
    If imgRequisitos.Tag = 0 Then
       MsgBox "Este Expediente no cumple con los requisitos para su Aprobación!", vbExclamation
       Exit Sub
    End If
    
    If imgTiempoPresentacion.Tag = 0 Then
       MsgBox "Este Expediente no cumple el Tiempo de Presentación para su Aprobación!", vbExclamation
       Exit Sub
    End If
    
    If imgExpedientesActivos.Tag = 0 Then
       MsgBox "Este Expediente No puede ser Aprobado por que ya fueron presentadas otras solicitudes de la misma persona (Revise DUPLICADOS!)", vbExclamation
       Exit Sub
    End If
End If


If Len(txtResolucionNotas.Text) < 10 Then
   MsgBox "Indique una nota válida para la resolución!", vbExclamation
   Exit Sub
End If

strSQL = "select NUMERO_RESOLUTORES from FSL_Comites" _
        & " where cod_Comite = '" & SIFGlobal.fxCodText(cboComite.Text) & "'"
Call OpenRecordSet(rs, strSQL)
    vNumResolutores = rs!NUMERO_RESOLUTORES
rs.Close

y = 0
With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked And Mid(.Item(i).SubItems(1), 1, 1) = "V" Then
            y = y + 1
      End If
   Next i
End With

If y < vNumResolutores Then
   MsgBox "Debe de indicar al menos (" & vNumResolutores & ") miembros del comité VALIDADOS! que den la resolución!", vbExclamation
   Exit Sub
End If



Me.MousePointer = vbHourglass


strSQL = "update FSL_EXPEDIENTES set RESOLUCION_NOTAS = '" & txtResolucionNotas.Text _
       & "',RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "', RESOLUCION_FECHA = getdate()" _
       & " ,RESOLUCION_USUARIO = '" & glogon.Usuario & "',ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "' where COD_EXPEDIENTE = " & txtExpediente.Text
Call ConectionExecute(strSQL)


strSQL = "delete FSL_EXPEDIENTE_COMITE WHERE COD_EXPEDIENTE = " & txtExpediente.Text
Call ConectionExecute(strSQL)


With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked Then
            strSQL = "INSERT FSL_EXPEDIENTE_COMITE(COD_EXPEDIENTE,COD_COMITE,CEDULA,ASIGNA_FECHA,ASIGNA_USUARIO,RESOLUCION_ESTADO)" _
                   & " values(" & txtExpediente & ",'" & SIFGlobal.fxCodText(cboComite.Text) & "','" & _
                   .Item(i).Tag & "',getdate(),'" & glogon.Usuario & "','" & Mid(cboResolucion.Text, 1, 1) & "')"
            Call ConectionExecute(strSQL)
      End If
   Next i
End With

Me.MousePointer = vbDefault

MsgBox "Expediente actualizado satisfactoriamente...", vbInformation

Call sbConsulta

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbValidaMiembro_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Valida"
     Call sbMiembroValida
  Case "Cerrar"
     fraValidaMiembro.Visible = False
     txtMiembroClave.Text = ""
End Select


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "CEDULA"
    gBusquedas.Columna = "CEDULA"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
txtNombre.Text = fxNombre(txtCedula)

End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then Call sbConsulta

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_EXPEDIENTE,CEDULA from FSL_EXPEDIENTES"
    gBusquedas.Orden = "COD_EXPEDIENTE"
    gBusquedas.Columna = "COD_EXPEDIENTE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtExpediente.Text = gBusquedas.Resultado
    txtCedula.Text = gBusquedas.Resultado2
    If Trim(txtExpediente.Text) <> "" Then Call sbConsulta
End If

End Sub

Private Sub txtExpediente_LostFocus()
If IsNumeric(txtExpediente.Text) Then
  Call sbConsulta
End If
End Sub



Private Function fxExpedienteConsecutivo(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(max(cod_Expediente),0) as 'Ultimo'" _
       & " from FSL_Expedientes where Cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)
    fxExpedienteConsecutivo = rs!Ultimo
rs.Close

End Function

Private Function fxPlanTipoDesembolso(pPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select TIPO_DESEMBOLSO" _
       & " from FSL_PLANES where cod_plan = '" & pPlan & "'"
Call OpenRecordSet(rs, strSQL)
    fxPlanTipoDesembolso = rs!TIPO_DESEMBOLSO
rs.Close

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoDesembolso As String

On Error GoTo vError


If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar este trámite porque no se encuentra pendiente", vbExclamation
    Exit Sub
End If

vTipoDesembolso = fxPlanTipoDesembolso(SIFGlobal.fxCodText(cboTipo.Text))


If Not vEdita Then
   strSQL = "insert FSL_EXPEDIENTES(COD_EXPEDIENTE,CEDULA, COD_PLAN, COD_CAUSA,COD_COMITE,COD_ENFERMEDAD,ESTADO,RESOLUCION_ESTADO" _
          & ", PRESENTA_CEDULA, PRESENTA_NOMBRE, PRESENTA_NOTAS, REFERENCIA_DOCUMENTO, REFERENCIA_NUMERO" _
          & ", ENFERMEDAD_FECHA,ENFERMEDAD_USUARIO, ENFERMEDAD_NOTAS, FECHA_ESTABLECE_CAUSA, NOTAS" _
          & ", TOTAL_DISPONIBLE, TOTAL_APLICADO, TOTAL_SOBRANTE, REGISTRO_FECHA, REGISTRO_USUARIO" _
          & ", TIPO_DESEMBOLSO)" _
          & " VALUES(dbo.fxFSL_ExpedienteConsecutivo(),'" & txtCedula.Text & "','" & SIFGlobal.fxCodText(cboTipo.Text) _
          & "','" & SIFGlobal.fxCodText(cboCausa.Text) & "','" & SIFGlobal.fxCodText(cboComite.Text) _
          & "','" & SIFGlobal.fxCodText(cboEnfermedad.Text) & "','P','P','" & txtPresentaCedula.Text _
          & "','" & txtPresentaNombre.Text & "','" & txtPresentaNotas.Text & "','" & cboRefTipoDoc.Text _
          & "','" & txtRefNumero.Text & "','" & Format(dtpEnfermedad.Value, "yyyy/mm/dd") _
          & "','" & glogon.Usuario & "','" & txtEnfermedadNotas.Text & "','" & Format(dtpRefFecha.Value, "yyyy/mm/dd") _
          & "','" & txtNotas.Text & "',0,0,0,getdate(),'" & glogon.Usuario & "','" & vTipoDesembolso & "')"
   
   Call ConectionExecute(strSQL)

   txtExpediente.Text = fxExpedienteConsecutivo(txtCedula.Text)
  

Else

  strSQL = "update FSL_EXPEDIENTES set COD_PLAN = '" & SIFGlobal.fxCodText(cboTipo.Text) _
         & "',COD_CAUSA = '" & SIFGlobal.fxCodText(cboCausa.Text) & "' ,COD_COMITE ='" & SIFGlobal.fxCodText(cboComite.Text) _
         & "',COD_ENFERMEDAD = '" & SIFGlobal.fxCodText(cboEnfermedad.Text) _
         & "',notas = '" & txtNotas.Text & "',PRESENTA_CEDULA = '" & txtPresentaCedula.Text & "', PRESENTA_NOMBRE = '" _
         & txtPresentaNombre.Text & "', REFERENCIA_DOCUMENTO = '" & cboRefTipoDoc.Text & "', REFERENCIA_NUMERO = '" _
         & txtRefNumero.Text & "', PRESENTA_NOTAS = '" & txtPresentaNotas.Text & "', FECHA_ESTABLECE_CAUSA = '" _
         & Format(dtpRefFecha.Value, "yyyy/mm/dd") & "', ENFERMEDAD_FECHA = '" & Format(dtpEnfermedad.Value, "yyyy/mm/dd") _
         & "',ENFERMEDAD_NOTAS = '" & txtEnfermedadNotas.Text & "', MODIFICA_USUARIO = '" & glogon.Usuario _
         & "', MODIFICA_FECHA = GETDATE(), TIPO_DESEMBOLSO = '" & vTipoDesembolso _
         & "' where COD_EXPEDIENTE = " & txtExpediente.Text
  Call ConectionExecute(strSQL)

End If

'Actualiza Requisitos
strSQL = "exec spFSL_ExpedienteRequisitos " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Actualiza Calculos de Creditos (FOSOL)
strSQL = "exec spFSL_ExpedienteOperaciones " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call sbToolBar(tlb, "activo")
Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbAplicarFosol()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset, pTipoDoc As String, pNumDoc As String

On Error GoTo vError

i = MsgBox("Esta seguro que aplicar los calculos del FOSOL?", vbYesNo)
If i = vbNo Then Exit Sub


Me.MousePointer = vbHourglass

strSQL = "exec spFSL_AplicacionFosol " & txtExpediente.Text & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
  pTipoDoc = rs!Tipo_Documento
  pNumDoc = rs!Numero_Documento
rs.Close

Me.MousePointer = vbDefault
MsgBox "Aplicación realizada satisfactoriamente..!", vbInformation

Call sbImprimeRecibo(pNumDoc, pTipoDoc)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

GLOBALES.gTag = txtExpediente.Text

Select Case Button.Key
 Case "Gestiones"
    Call sbFormsCall("frmFSL_ExpedienteGestiones", 1, , , False, Me)
 
 Case "Apelacion"
    Call sbFormsCall("frmFSL_ExpedienteApelaciones", 1, , , False, Me)
 
 Case "Aplicar"
    Call sbAplicarFosol
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbConsultaExterna(xTramTemp As String)
 txtExpediente.Text = xTramTemp
 If Trim(txtExpediente) <> "" Then
    Call txtExpediente_KeyDown(vbKeyReturn, 0)
 End If

End Sub


Private Sub sbMiembroValida()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

 

'Verifica Usuario / Cifrado Actual
strSQL = "exec spSEGLogon '" & txtMiembroUsuario.Text & "','" & SIFGlobal.fxStringCifrado(txtMiembroClave.Text) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   MsgBox "No fue posible validar al usuario.: Verifique su contraseña!", vbExclamation
   txtMiembroClave.SetFocus
Else
 
 With lswComite.ListItems
   For i = 1 To .Count
     If .Item(i).Tag = lblMiembro.Tag Then
        .Item(i).SubItems(1) = "Validado"
        fraValidaMiembro.Visible = False
     End If
   Next i
 End With
End If
rs.Close

txtMiembroClave.Text = ""

End Sub

Private Sub txtMiembroClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbMiembroValida
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaCedula.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "NOMBRE"
    gBusquedas.Columna = "NOMBRE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "CEDULA"
    gBusquedas.Columna = "CEDULA"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtPresentaCedula.Text = gBusquedas.Resultado
    txtPresentaNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaCedula_LostFocus()
Dim pNombre As String

pNombre = fxNombre(txtPresentaCedula.Text)

If Trim(pNombre) <> "" Then
   txtPresentaNombre.Text = pNombre
End If

End Sub

Private Sub txtPresentaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPresentaNotas.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select CEDULA,NOMBRE from SOCIOS"
    gBusquedas.Orden = "NOMBRE"
    gBusquedas.Columna = "NOMBRE"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtPresentaCedula.Text = gBusquedas.Resultado
    txtPresentaNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPresentaNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Then cboRefTipoDoc.SetFocus

End Sub

Private Sub txtRefNumero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpRefFecha.SetFocus

End Sub

