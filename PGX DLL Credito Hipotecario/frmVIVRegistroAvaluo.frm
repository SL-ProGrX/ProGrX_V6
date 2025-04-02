VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmVivRegistroAvaluo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de avalúo de la propiedad"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   13905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   9375
      _Version        =   1441793
      _ExtentX        =   16536
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Información del Crédito"
      ForeColor       =   16711680
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8640
         Top             =   -480
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
               Picture         =   "frmVIVRegistroAvaluo.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVIVRegistroAvaluo.frx":6862
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   480
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
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   840
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
      Begin XtremeSuiteControls.FlatEdit txtExpediente 
         Height          =   315
         Left            =   7080
         TabIndex        =   8
         Top             =   480
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
         Height          =   315
         Left            =   4320
         TabIndex        =   9
         Top             =   840
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
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
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Operación"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   10
         Top             =   480
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Expediente"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   13905
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   4260
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   4455
         TabIndex        =   2
         Top             =   30
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   582
         ButtonWidth     =   2117
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cobertura"
               Key             =   "cobertura"
               Object.ToolTipText     =   "Cobertura de la Propiedad"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Crédito"
               Key             =   "montocredito"
               Object.ToolTipText     =   "Corrección monto del crédito"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   9375
      _Version        =   1441793
      _ExtentX        =   16536
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Información de la Garantía"
      ForeColor       =   16711680
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
      Begin XtremeSuiteControls.FlatEdit txtDistrito 
         Height          =   315
         Left            =   7080
         TabIndex        =   13
         Top             =   1440
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
      Begin XtremeSuiteControls.FlatEdit txtProvincia 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   1440
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
      Begin XtremeSuiteControls.FlatEdit txtIngenieroNombre 
         Height          =   315
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
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
      Begin XtremeSuiteControls.FlatEdit txtIngenieroId 
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
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
      Begin XtremeSuiteControls.FlatEdit txtZona 
         Height          =   315
         Left            =   7080
         TabIndex        =   17
         Top             =   360
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
      Begin XtremeSuiteControls.FlatEdit txtFinca 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Top             =   360
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
      Begin XtremeSuiteControls.FlatEdit txtArea 
         Height          =   315
         Left            =   7080
         TabIndex        =   19
         Top             =   720
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
      Begin XtremeSuiteControls.FlatEdit txtCanton 
         Height          =   315
         Left            =   4320
         TabIndex        =   20
         Top             =   1440
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.FlatEdit txtPlano 
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   720
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
      Begin XtremeSuiteControls.FlatEdit txtDireccion 
         Height          =   675
         Left            =   2160
         TabIndex        =   22
         Top             =   1800
         Width           =   7095
         _Version        =   1441793
         _ExtentX        =   12515
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Finca"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Plano Catastro"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   8
         Left            =   5160
         TabIndex        =   26
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Zona"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   9
         Left            =   5160
         TabIndex        =   25
         Top             =   720
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Área (m2)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ingeniero"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Dirección"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2655
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   13695
      _Version        =   1441793
      _ExtentX        =   24156
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Registro Informativo del Avalúo"
      ForeColor       =   16711680
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
      Begin XtremeSuiteControls.RadioButton rbPoliza 
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Personal"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   315
         Left            =   2160
         TabIndex        =   32
         Top             =   2280
         Width           =   2175
         _Version        =   1441793
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
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtValorConstruccion 
         Height          =   315
         Left            =   2160
         TabIndex        =   33
         Top             =   1200
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtValorTerreno 
         Height          =   315
         Left            =   2160
         TabIndex        =   34
         Top             =   840
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaInspeccion 
         Height          =   315
         Left            =   2160
         TabIndex        =   35
         Top             =   360
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   1995
         Left            =   4680
         TabIndex        =   36
         Top             =   600
         Width           =   8895
         _Version        =   1441793
         _ExtentX        =   15690
         _ExtentY        =   3519
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
      Begin XtremeSuiteControls.RadioButton rbPoliza 
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comercial"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtPrima 
         Height          =   315
         Left            =   2160
         TabIndex        =   50
         Top             =   1920
         Width           =   2175
         _Version        =   1441793
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
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   51
         Top             =   1920
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto de Prima"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Inspección"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   42
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   41
         Top             =   840
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor del terreno"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   16
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor contrucción"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   39
         Top             =   2280
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Valor total inmueble"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   18
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Póliza"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1575
      Index           =   3
      Left            =   9720
      TabIndex        =   29
      Top             =   1080
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Información del Ingeniero"
      ForeColor       =   16711680
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
      Begin XtremeSuiteControls.PushButton btnIngCambios 
         Height          =   330
         Index           =   0
         Left            =   2640
         TabIndex        =   48
         Top             =   600
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Cambiar"
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
      End
      Begin XtremeSuiteControls.FlatEdit txtAvaluo 
         Height          =   315
         Left            =   360
         TabIndex        =   44
         Top             =   600
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtViaticos 
         Height          =   315
         Left            =   360
         TabIndex        =   46
         Top             =   1200
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnIngCambios 
         Height          =   330
         Index           =   1
         Left            =   2640
         TabIndex        =   49
         Top             =   1200
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Cambiar"
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
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   47
         Top             =   960
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Viáticos/Kilometraje"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   45
         ToolTipText     =   "Valor del Avalúo"
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto por Avalúo"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1575
      Index           =   4
      Left            =   9720
      TabIndex        =   30
      Top             =   2880
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Información del Abogado"
      ForeColor       =   16711680
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
      Begin XtremeSuiteControls.FlatEdit txtA_Honorarios 
         Height          =   315
         Left            =   360
         TabIndex        =   52
         Top             =   600
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtA_GastosLegales 
         Height          =   315
         Left            =   360
         TabIndex        =   54
         Top             =   1200
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
         Text            =   "0.00"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnIngCambios 
         Height          =   330
         Index           =   2
         Left            =   2640
         TabIndex        =   56
         Top             =   600
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Cambiar"
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
      End
      Begin XtremeSuiteControls.PushButton btnIngCambios 
         Height          =   330
         Index           =   3
         Left            =   2640
         TabIndex        =   57
         Top             =   1200
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Cambiar"
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
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   19
         Left            =   360
         TabIndex        =   55
         Top             =   960
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Gastos Legales"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   53
         ToolTipText     =   "Valor del Avalúo"
         Top             =   360
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto Honorarios"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmVivRegistroAvaluo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cambioDatos As Boolean
Public vNumOperacion As String
Public vIdGarantia As Long
Public vIdcontacto As Long
Private vEditar As Boolean

Private Sub dtpFechaInspeccion_Click()
m_cambioDatos = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.ActiveControl.Name = "TxtObservaciones" Then Exit Sub

If (KeyCode = vbKeyReturn) Then
    Call gsbPulsarTecla(vbKeyTab)
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo vError

Select Case Me.ActiveControl.Name
Case "txtValorTerreno", "txtViaticos"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtValorTerreno.Text), KeyAscii)
Case "txtValorConstruccion"
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtValorConstruccion.Text), KeyAscii)

End Select

salir:
    Exit Sub
vError:
    MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation
    Resume salir
End Sub

Private Sub Form_Load()
 vEditar = True
 
 vNumOperacion = GLOBALES.gTag
 vIdGarantia = GLOBALES.gTag2
 vIdcontacto = GLOBALES.gTag3
 
txtViaticos.Text = Format(0, "Standard")
txtValorConstruccion.Text = Format(0, "Standard")
txtValorTerreno.Text = Format(0, "Standard")

txtTotal.Text = Format(0, "Standard")
dtpFechaInspeccion.Value = fxFechaServidor 'Format(fxFechaServidor, "DD/MM/YYYY") ' ObjConsultar.fxFechaServer

Call sbToolBarIconos(tlbPrincipal, False)

txtOperacion.Text = vNumOperacion

 

Call sbTraerInformacionOperacion
m_cambioDatos = False
End Sub

Private Sub sbTraerInformacionOperacion()
On Error GoTo vError
 
If ObjConsultar.fxTraerOperacionXIdGarantiaIng(vNumOperacion, vIdGarantia) Then
    With glogon.Recordset.Fields
    
    
    vIdcontacto = !IdContacto
    
    txtOperacion.Text = vNumOperacion
    txtCedula.Text = Trim(!cedula)
     
    txtNombre.Text = (!Nombre)
    txtExpediente.Text = IIf(IsNull(!Expediente), "", Trim(!Expediente))
    txtFinca.Text = Trim(!NumeroFinca)
    txtPlano.Text = IIf(IsNull(!NumPlanoCatastro), "", Trim(!NumPlanoCatastro))
    
    txtIngenieroId.Text = !IdContacto & ""
    txtIngenieroNombre.Text = Trim(!NombreProfesional)
    txtZona.Text = Trim(!DescZona)
    txtArea.Text = Trim(!AreaFinca)
    txtProvincia.Text = Trim(!PROVINCIA)
    txtCanton.Text = Trim(!Canton & "")
    txtDistrito.Text = Trim(!Distrito & "")
    txtDireccion.Text = Trim(!Direccion & "")
    
    If Not IsNull(!FechaInspeccion) Then
        dtpFechaInspeccion.Value = Format(!FechaInspeccion, "dd-mm-yyyy")
        vEditar = False
    Else
        vEditar = True
    End If
    
    txtValorTerreno.Text = IIf(IsNull(!ValorTerreno), Format(0, "Standard"), Format(!ValorTerreno, "Standard"))
    txtValorConstruccion.Text = IIf(IsNull(!ValorConstruccion), Format(0, "Standard"), Format(!ValorConstruccion, "Standard"))

    txtObservaciones.Text = IIf(IsNull(!ObservacionAvaluo), "", Trim(!ObservacionAvaluo))
    txtViaticos.Text = IIf(IsNull(!Viaticos), Format(0, "Standard"), Format(!Viaticos, "Standard"))
    txtViaticos.Enabled = True
    txtViaticos.BackColor = &H80000005
    
    If !Tipo_Poliza = "P" Then
        rbPoliza.Item(0).Value = True
    Else
        rbPoliza.Item(1).Value = True
    End If
    
    End With
End If

salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError
    
Select Case Button.Key

    Case "cobertura"
        gOperacion = txtOperacion.Text
        Call sbSIFForms("frmVivCoberturas", 1, , , False)
    Case "montocredito"
        gOperacion = txtOperacion.Text
        Call sbSIFForms("frmVivCorregirMontoCredito", 1, , , False)
    
End Select
    

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
    Select Case Button.Key
        Case "nuevo"
            
        Case "editar"

        Case "borrar"
        
        Case "guardar"
            Call sbAgregar
            
        Case "deshacer"

    End Select
    
salir:
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    Resume salir
End Sub

Private Function fxValidaRegistroAvaluo(ByVal pIdGarantia As Long) As Boolean

On Error GoTo vError

fxValidaRegistroAvaluo = False
                
glogon.strSQL = "SELECT G.ValorConstruccion, G.ValorTerreno" & _
                " FROM  ViviendaGarantia AS G" & _
                " where G.IdGarantia = " & pIdGarantia
          
                       
If execSql(glogon.strSQL, True) Then
    fxValidaRegistroAvaluo = IIf(IsNull(glogon.Recordset.Fields!ValorConstruccion), False, True)
End If
Exit Function

vError:
MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Function

Private Sub sbAgregar()
On Error GoTo vError


If m_cambioDatos = False Then Exit Sub
If fxValidaDatos() = False Then Exit Sub

If (MsgBox("¿Desea guardar la información digitada.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
Me.MousePointer = vbHourglass

If ObjAgregar.fxRegistroAvaluo(gParametros(1), gParametros(2), gParametros(3), gParametros(4), _
                                                   gParametros(5), gParametros(6), gParametros(7), gParametros(8), gParametros(9), gParametros(10)) Then
    m_cambioDatos = False
    
    Call Bitacora("APLICA", "Registro avaluo Garantia Vivienda: " & gParametros(1) & " Contacto: " & gParametros(2))
    
    MsgBox "Información fue registrada corretamente.", vbInformation
     vEditar = False
End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub
'---------------------Guardar informacion de duenos segun numero de garantia--------------------
Private Function fxValidaDatos() As Boolean
On Error GoTo vError

fxValidaDatos = False

ReDim gParametros(1 To 10)
If fxValidaRegistroAvaluo(vIdGarantia) Then
   Me.MousePointer = vbDefault
    MsgBox ("La información de avaluo no puede ser modificada, ya fue registrado")
    Exit Function
End If

gParametros(1) = vIdGarantia
gParametros(2) = vIdcontacto
gParametros(3) = Format(dtpFechaInspeccion.Value, "yyyy/mm/dd")

If Not IsNumeric(txtValorTerreno.Text) Then
    gParametros(4) = 0
Else
    gParametros(4) = CCur(txtValorTerreno.Text)
End If
If Not IsNumeric(txtValorConstruccion.Text) Then
    gParametros(5) = 0
Else
    gParametros(5) = CCur(txtValorConstruccion.Text)
End If

gParametros(6) = IIf((Len(txtObservaciones.Text) = 0), ObjNull.NullString, Trim(txtObservaciones.Text))
gParametros(7) = glogon.Usuario
gParametros(8) = "1900/01/01"

If Not IsNumeric(txtViaticos.Text) Then
    gParametros(9) = 0
Else
    gParametros(9) = CCur(txtViaticos.Text)
End If

If rbPoliza.Item(0).Value Then
    gParametros(10) = "P"
Else
    gParametros(10) = "C"
End If

fxValidaDatos = True

salir:
    Exit Function
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub TxtObservaciones_Change()
m_cambioDatos = True
End Sub

Private Sub txtValorConstruccion_Change()
m_cambioDatos = True

If Not IsNumeric(txtValorTerreno.Text) Then Exit Sub
If Not IsNumeric(txtValorConstruccion.Text) Then Exit Sub

txtTotal.Text = (CCur(txtValorConstruccion.Text) + CCur(txtValorTerreno.Text))

If Val(txtTotal.Text) = 0 Then Exit Sub
txtTotal.Text = Format(txtTotal.Text, "Standard")
End Sub


Private Sub txtValorConstruccion_GotFocus()
txtValorConstruccion.SelStart = 0
txtValorConstruccion.SelLength = Len(txtValorConstruccion.Text)
If Val(txtValorConstruccion.Text) = 0 Then
 txtValorConstruccion.Text = 0
Else
    txtValorConstruccion.Text = CCur(txtValorConstruccion.Text)
End If


End Sub

Private Sub txtValorConstruccion_LostFocus()
If Val(txtValorConstruccion.Text) = 0 Then
txtValorConstruccion.Text = Format(0, "Standard")
Else
txtValorConstruccion.Text = Format(txtValorConstruccion.Text, "Standard")
End If
End Sub

Private Sub txtValorTerreno_Change()
m_cambioDatos = True
If Not IsNumeric(txtValorTerreno.Text) Then Exit Sub
If Not IsNumeric(txtValorConstruccion.Text) Then Exit Sub

txtTotal.Text = CCur(txtValorConstruccion.Text) + CCur(txtValorTerreno.Text)


If Len(txtTotal.Text) = 0 Then Exit Sub
txtTotal.Text = Format(txtTotal.Text, "Standard")
End Sub


Private Sub txtValorTerreno_GotFocus()
txtValorTerreno.SelStart = 0
txtValorTerreno.SelLength = Len(txtValorTerreno.Text)
If Val(txtValorTerreno.Text) = 0 Then Exit Sub
txtValorTerreno.Text = CCur(txtValorTerreno.Text)
End Sub

Private Sub txtValorTerreno_LostFocus()

If Val(txtValorTerreno.Text) = 0 Then
    txtValorTerreno.Text = Format(0, "Standard")
Else
    txtValorTerreno.Text = Format(txtValorTerreno.Text, "Standard")
End If


End Sub

Private Sub txtViaticos_Change()
m_cambioDatos = True
End Sub

Private Sub txtViaticos_GotFocus()
txtViaticos.SelStart = 0
txtViaticos.SelLength = Len(txtViaticos.Text)
If Val(txtViaticos.Text) = 0 Then Exit Sub
txtViaticos.Text = CCur(txtViaticos.Text)

End Sub

Private Sub txtViaticos_LostFocus()
If Val(txtViaticos.Text) = 0 Then
    txtViaticos.Text = Format(0, "Standard")
Else
    txtViaticos.Text = Format(txtViaticos.Text, "Standard")
End If

End Sub
