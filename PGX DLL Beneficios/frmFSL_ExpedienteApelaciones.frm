VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmFSL_ExpedienteApelaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Apelaciones"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      _Version        =   1441793
      _ExtentX        =   20558
      _ExtentY        =   11880
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
      SelectedItem    =   1
      Item(0).Caption =   "Histórico"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vgApelaciones"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "scTitulos(0)"
      Item(1).Control(1)=   "txtPresentaNombre"
      Item(1).Control(2)=   "txtPresentaCedula"
      Item(1).Control(3)=   "txtPresentaNotas"
      Item(1).Control(4)=   "label5(2)"
      Item(1).Control(5)=   "label5(1)"
      Item(1).Control(6)=   "label5(0)"
      Item(1).Control(7)=   "label5(3)"
      Item(1).Control(8)=   "cboApelacion"
      Item(1).Control(9)=   "cmdAplicar"
      Item(2).Caption =   "Resolución"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "fraValidaMiembro"
      Item(2).Control(1)=   "txtResolucionNotas"
      Item(2).Control(2)=   "cboResolucion"
      Item(2).Control(3)=   "lswComite"
      Item(2).Control(4)=   "Label3(11)"
      Item(2).Control(5)=   "Label3(12)"
      Item(2).Control(6)=   "Label3(13)"
      Item(2).Control(7)=   "tlbResolucion"
      Begin XtremeSuiteControls.ListView lswComite 
         Height          =   4095
         Left            =   -64720
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin VB.Frame fraValidaMiembro 
         Caption         =   "Validación del Miembro de Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -67720
         TabIndex        =   19
         Top             =   1680
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
            TabIndex        =   21
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
            TabIndex        =   20
            ToolTipText     =   "Número de Tramite"
            Top             =   720
            Width           =   2175
         End
         Begin MSComctlLib.Toolbar tlbValidaMiembro 
            Height          =   330
            Left            =   3000
            TabIndex        =   22
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
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Usuario .: "
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Clave .: "
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   24
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   5295
         End
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   615
         Left            =   5880
         TabIndex        =   18
         Top             =   3720
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Picture         =   "frmFSL_ExpedienteApelaciones.frx":0000
         ImageAlignment  =   4
      End
      Begin FPSpreadADO.fpSpread vgApelaciones 
         Height          =   6135
         Left            =   -69880
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   11295
         _Version        =   524288
         _ExtentX        =   19923
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
         MaxCols         =   7
         SpreadDesigner  =   "frmFSL_ExpedienteApelaciones.frx":0727
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPresentaNombre 
         Height          =   330
         Left            =   4200
         TabIndex        =   10
         Top             =   1200
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
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
      Begin XtremeSuiteControls.FlatEdit txtPresentaCedula 
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Top             =   1200
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
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
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
      Begin XtremeSuiteControls.ComboBox cboApelacion 
         Height          =   330
         Left            =   1800
         TabIndex        =   17
         Top             =   3000
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
      Begin MSComctlLib.Toolbar tlbResolucion 
         Height          =   330
         Left            =   -62920
         TabIndex        =   29
         Top             =   5400
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
      Begin XtremeSuiteControls.ComboBox cboResolucion 
         Height          =   330
         Left            =   -66400
         TabIndex        =   30
         Top             =   5400
         Visible         =   0   'False
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5953
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
      Begin XtremeSuiteControls.FlatEdit txtResolucionNotas 
         Height          =   4050
         Left            =   -69760
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   7144
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
         Alignment       =   1  'Right Justify
         Caption         =   " Resolución.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -68200
         TabIndex        =   28
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Miembros del Comité.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -64720
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Notas de la Resolución.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -69760
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin XtremeSuiteControls.Label label5 
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Tipo de apelación "
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
         Left            =   1800
         TabIndex        =   15
         Top             =   960
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
      Begin XtremeSuiteControls.Label label5 
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   14
         Top             =   960
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
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Notas de la Apelación"
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
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   12975
         _Version        =   1441793
         _ExtentX        =   22886
         _ExtentY        =   873
         _StockProps     =   14
         Caption         =   "Datos de la persona que presenta la apelación:"
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
   Begin MSComctlLib.ImageList imgLista 
      Left            =   12720
      Top             =   360
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
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":0E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":76B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_ExpedienteApelaciones.frx":77B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   120
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
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   120
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
      TabIndex        =   4
      Top             =   840
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   4440
      TabIndex        =   5
      Top             =   840
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Top             =   120
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   330
      Left            =   720
      TabIndex        =   6
      Top             =   840
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
End
Attribute VB_Name = "frmFSL_ExpedienteApelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mComite As String, mLinea As Integer

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If txtEstado.Tag <> "R" Then
   MsgBox "El expediente no se encuentra RECHAZADO para poder registrar una apelación!", vbExclamation
   Exit Sub
End If

If mLinea > 0 Then
   MsgBox "Ya se encuentra registrada una apelación (Pendiente de Resolución) a este expediente, verifique!", vbExclamation
   Exit Sub
End If

On Error GoTo vError



strSQL = "exec spFSL_ApelacionRegistra " & txtExpediente.Text & ",'" & cboApelacion.ItemData(cboApelacion.ListIndex) & "','" _
       & txtPresentaCedula.Text & "','" _
       & txtPresentaNombre.Text & "','" & txtPresentaNotas.Text & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Apelación registrada satisfactoriamente!", vbInformation
Call sbInicializa

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()

vModulo = 7

txtExpediente.Text = GLOBALES.gTag

With lswComite.ColumnHeaders
    .Clear
    .Add , , "Comité", 5000
    .Add , , "", 1200
End With

Call sbInicializa

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass


tcMain.Item(0).Selected = True


strSQL = "select rtrim(cod_apelacion) as 'IdX', rtrim(DESCRIPCION) as 'ItmX' from FSL_TIPOS_APELACIONES WHERE ACTIVA = 1"
Call sbCbo_Llena_New(cboApelacion, strSQL, False, True)

cboResolucion.Clear
cboResolucion.AddItem "Aprobado"
cboResolucion.AddItem "Rechazado"
cboResolucion.Text = "Aprobado"



strSQL = "select Soc.Nombre,Ex.*" _
       & " from FSL_Expedientes Ex inner join Socios Soc on Ex.cedula = Soc.Cedula" _
       & " Where Ex.Cod_Expediente = " & txtExpediente.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Or Not rs.BOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre

  
  txtPresentaCedula.Text = rs!Presenta_Cedula & ""
  txtPresentaNombre.Text = rs!PRESENTA_NOMBRE & ""
  txtPresentaNotas.Text = ""
  
  mComite = rs!cod_Comite
 txtEstado.Tag = rs!Estado
  Select Case rs!Estado
   Case "P" 'Pendiente
        txtEstado.Text = "PENDIENTE"
    Case "A" 'Aprobado
        txtEstado.Text = "APROBADO"
    Case "R" 'Rechazado
        txtEstado.Text = "RECHAZADO"
    Case "X" 'Aplicado
        txtEstado.Text = "APLICADO"
  End Select

End If
rs.Close

'Linea de Apelación Pendiente
mLinea = 0

strSQL = "select isnull( max(Linea),0) as 'Linea'" _
       & " from FSL_EXPEDIENTES_APELACIONES" _
       & " Where cod_Expediente = " & txtExpediente.Text & " and resolucion = 'P'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Or Not rs.BOF Then
  mLinea = rs!Linea
End If
rs.Close

'Histórico
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


'Resolucion
strSQL = "Select Cm.Cedula,Cm.Nombre,isnull(Ec.Asigna_Usuario,'No!') as 'ASIGNADO'" _
       & " from FSL_EXPEDIENTES Ex" _
       & "  inner join FSL_COMITES_MIEMBROS Cm on Ex.COD_COMITE = Cm.COD_COMITE" _
       & "   left join FSL_EXPEDIENTE_COMITE Ec on Ex.COD_EXPEDIENTE = Ec.COD_EXPEDIENTE and Ex.COD_COMITE = Ec.COD_COMITE" _
       & "         and Cm.Cedula = Ex.Cedula" _
       & " where Ex.cod_Expediente = " & txtExpediente.Text & " and Cm.Activo = 1"
 Call OpenRecordSet(rs, strSQL)
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



Me.MousePointer = vbDefault

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
           & " where cedula = '" & .Tag & "' and cod_comite = '" & mComite & "'"
     Call OpenRecordSet(rs, strSQL)
         txtMiembroUsuario.Text = rs!Usuario_Vinculado
     rs.Close
     
     txtMiembroClave.SetFocus

  End If
  
End With
End Sub

Private Sub tlbResolucion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim y As Integer, vNumResolutores As Integer


If txtEstado.Tag = "X" Then
   MsgBox "Este Expediente se encuentra aplicado no se puede cambiar la resolución!", vbExclamation
   Exit Sub
End If

If mLinea = 0 Then
   MsgBox "No existe ninguna apelación pendiente de resolución", vbExclamation
   Exit Sub
End If


If Len(txtResolucionNotas.Text) < 10 Then
   MsgBox "Indique una nota válida para la resolución!", vbExclamation
   Exit Sub
End If

strSQL = "select NUMERO_RESOLUTORES from FSL_Comites" _
        & " where cod_Comite = '" & mComite & "'"
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

On Error GoTo vError
Me.MousePointer = vbHourglass


strSQL = "update FSL_EXPEDIENTES set RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "', ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "' where COD_EXPEDIENTE = " & txtExpediente.Text
Call ConectionExecute(strSQL)

strSQL = "update FSL_EXPEDIENTES set RESOLUCION_ESTADO = '" & Mid(cboResolucion.Text, 1, 1) _
       & "', ESTADO = '" & Mid(cboResolucion.Text, 1, 1) & "' where COD_EXPEDIENTE = " & txtExpediente.Text
Call ConectionExecute(strSQL)


strSQL = "update FSL_EXPEDIENTES_APELACIONES set RESOLUCION_NOTAS = '" & txtResolucionNotas.Text _
       & "',RESOLUCION = '" & Mid(cboResolucion.Text, 1, 1) _
       & "',RESOLUCION_FECHA = getdate(), RESOLUCION_USUARIO = '" & glogon.Usuario _
       & "' where COD_EXPEDIENTE = " & txtExpediente.Text & " and Linea = " & mLinea
Call ConectionExecute(strSQL)


strSQL = "delete FSL_EXPEDIENTES_APELACIONES_COMITE WHERE COD_EXPEDIENTE = " & txtExpediente.Text
Call ConectionExecute(strSQL)


With lswComite.ListItems
   For i = 1 To .Count
      If .Item(i).Checked Then
            strSQL = "INSERT FSL_EXPEDIENTES_APELACIONES_COMITE(LINEA,COD_EXPEDIENTE,COD_COMITE,CEDULA,ASIGNA_FECHA,ASIGNA_USUARIO)" _
                   & " values(" & mLinea & "," & txtExpediente & ",'" & mComite & "','" & _
                   .Item(i).Tag & "',getdate(),'" & glogon.Usuario & "')"
            Call ConectionExecute(strSQL)
      End If
   Next i
End With

Me.MousePointer = vbDefault

MsgBox "Expediente actualizado satisfactoriamente...", vbInformation

Call sbInicializa

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
