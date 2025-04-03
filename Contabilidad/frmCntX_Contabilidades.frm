VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCntX_Contabilidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Contabilidades"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   HelpContextID   =   4
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtNivel8 
      Height          =   372
      Left            =   9120
      TabIndex        =   37
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel7 
      Height          =   372
      Left            =   7920
      TabIndex        =   35
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel6 
      Height          =   372
      Left            =   6720
      TabIndex        =   33
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel5 
      Height          =   372
      Left            =   5520
      TabIndex        =   31
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel4 
      Height          =   372
      Left            =   4320
      TabIndex        =   29
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel3 
      Height          =   372
      Left            =   3120
      TabIndex        =   27
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNivel2 
      Height          =   372
      Left            =   1920
      TabIndex        =   25
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.CheckBox chkExpCuentas 
      Height          =   252
      Left            =   360
      TabIndex        =   13
      Top             =   4800
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7429
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Catálogo Contable"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Contabilidades.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9660
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   4695
      NewRow1         =   0   'False
      Child2          =   "tlbUsuarios"
      MinHeight2      =   330
      Width2          =   300
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbUsuarios 
         Height          =   330
         Left            =   4890
         TabIndex        =   9
         Top             =   30
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   582
         ButtonWidth     =   2170
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Usuarios"
               Key             =   "Usuarios"
               Object.ToolTipText     =   "Usuarios asignados por Contabilidad"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.CheckBox chkExpAsientosGenerales 
      Height          =   252
      Left            =   360
      TabIndex        =   14
      Top             =   5040
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7429
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Asientos Generales"
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
   Begin XtremeSuiteControls.CheckBox chkExpMantenimiento 
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   5280
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7429
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Mantenimiento General"
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
   Begin XtremeSuiteControls.CheckBox chkExpDiferidos 
      Height          =   252
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Width           =   4212
      _Version        =   1572864
      _ExtentX        =   7429
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estructuras de Diferidos (Gastos / Ingresos)"
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
   Begin XtremeSuiteControls.CheckBox chkExpAreas 
      Height          =   252
      Left            =   4680
      TabIndex        =   17
      Top             =   4800
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Areas de Trabajo"
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
   Begin XtremeSuiteControls.CheckBox chkExpPlanFijos 
      Height          =   252
      Left            =   4680
      TabIndex        =   18
      Top             =   5040
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Plantillas de Asientos Fijos y Proyectados"
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
   Begin XtremeSuiteControls.CheckBox chkExpPlanRate 
      Height          =   252
      Left            =   4680
      TabIndex        =   19
      Top             =   5280
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Plantillas de Asientos Prorrateados"
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
   Begin XtremeSuiteControls.CheckBox chkExpPresupuesto 
      Height          =   252
      Left            =   4680
      TabIndex        =   20
      Top             =   5520
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Asignación Presupuestaria"
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
   Begin XtremeSuiteControls.FlatEdit txtMascara 
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   3840
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16954
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
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
   Begin XtremeSuiteControls.FlatEdit txtNivel1 
      Height          =   372
      Left            =   720
      TabIndex        =   23
      Top             =   3000
      Width           =   492
      _Version        =   1572864
      _ExtentX        =   868
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1800
      TabIndex        =   39
      Top             =   600
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
      BackColor       =   16777215
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   3000
      TabIndex        =   40
      Top             =   600
      Width           =   5295
      _Version        =   1572864
      _ExtentX        =   9334
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedulaJuridica 
      Height          =   312
      Left            =   1800
      TabIndex        =   41
      Top             =   1080
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTelCentral 
      Height          =   312
      Left            =   1800
      TabIndex        =   42
      Top             =   1440
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTelFax 
      Height          =   312
      Left            =   5760
      TabIndex        =   43
      Top             =   1440
      Width           =   2532
      _Version        =   1572864
      _ExtentX        =   4466
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmail 
      Height          =   312
      Left            =   1800
      TabIndex        =   44
      Top             =   1800
      Width           =   6492
      _Version        =   1572864
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
   Begin XtremeSuiteControls.FlatEdit txtContacto 
      Height          =   312
      Left            =   1800
      TabIndex        =   45
      Top             =   2160
      Width           =   6492
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ComboBox cboRazon 
      Height          =   312
      Left            =   5760
      TabIndex        =   46
      Top             =   1080
      Width           =   2532
      _Version        =   1572864
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
   Begin XtremeSuiteControls.CheckBox chkFiltra_Contabilidad 
      Height          =   255
      Left            =   360
      TabIndex        =   48
      Top             =   6360
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Módulos Contables"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra_Operaciones 
      Height          =   255
      Left            =   360
      TabIndex        =   49
      Top             =   6720
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Módulos Operaciones"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra_Inversiones 
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   6360
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Inversiones"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra_RRHH 
      Height          =   255
      Left            =   3600
      TabIndex        =   51
      Top             =   6720
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Recursos Humanos"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltra_Bancos 
      Height          =   255
      Left            =   6480
      TabIndex        =   52
      Top             =   6360
      Width           =   2655
      _Version        =   1572864
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Bancos"
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
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkConsolida 
      Height          =   255
      Left            =   360
      TabIndex        =   54
      Top             =   7680
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Esta Contabilidad es utilizada como consolidadora"
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
   Begin XtremeSuiteControls.ComboBox cboConsolida_Conta 
      Height          =   330
      Left            =   3240
      TabIndex        =   56
      Top             =   8160
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.ComboBox cboConsolida_Unidad 
      Height          =   330
      Left            =   3240
      TabIndex        =   58
      Top             =   8640
      Width           =   6255
      _Version        =   1572864
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.Label Label7 
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   57
      Top             =   8520
      Width           =   2775
      _Version        =   1572864
      _ExtentX        =   4895
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Indique la Unidad en la que se reporta la Contabilidad BASE "
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
   Begin XtremeSuiteControls.Label Label7 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   55
      Top             =   8040
      Width           =   2775
      _Version        =   1572864
      _ExtentX        =   4895
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Indique la Contabilidad BASE para Referencia de Consolidación"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   53
      Top             =   7200
      Width           =   9735
      _Version        =   1572864
      _ExtentX        =   17171
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Contabilidad de Consolidación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   47
      Top             =   6000
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16954
      _ExtentY        =   444
      _StockProps     =   14
      Caption         =   "Activar Control de Acceso a Cuentas por Roles"
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
   End
   Begin VB.Label Label2 
      Caption         =   "Email"
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
      Left            =   480
      TabIndex        =   38
      Top             =   1800
      Width           =   1092
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   7
      Left            =   8520
      TabIndex        =   36
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 8"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   6
      Left            =   7320
      TabIndex        =   34
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 7"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   5
      Left            =   6120
      TabIndex        =   32
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 6"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   4
      Left            =   4920
      TabIndex        =   30
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 5"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   3
      Left            =   3720
      TabIndex        =   28
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 4"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   2
      Left            =   2520
      TabIndex        =   26
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 3"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   1
      Left            =   1320
      TabIndex        =   24
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 2"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nivel 1"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblMascara 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3480
      Width           =   9615
      _Version        =   1572864
      _ExtentX        =   16960
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Muestra la máscara contable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   4440
      Width           =   9612
      _Version        =   1572864
      _ExtentX        =   16954
      _ExtentY        =   444
      _StockProps     =   14
      Caption         =   "Indique las Opciones Adicionales visibles en el Explorar Contable"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   252
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   9612
      _Version        =   1572864
      _ExtentX        =   16954
      _ExtentY        =   444
      _StockProps     =   14
      Caption         =   "Definición de Máscara Contable - Indique el número de caracteres por Nivel"
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
   End
   Begin VB.Label Label5 
      Caption         =   "Razón Social"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "Tel.Fax"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "Tel.Central"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "Contacto"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label C 
      Caption         =   "Ced.Jurídica"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "Contabilidad"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmCntX_Contabilidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vTipoBusca As String
Dim vScroll As Boolean
Dim vPaso As Boolean


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True


If Not IsNumeric(txtNivel1.Text) Or Not IsNumeric(txtNivel2.Text) Or Not IsNumeric(txtNivel2.Text) _
    Or Not IsNumeric(txtNivel4.Text) Or Not IsNumeric(txtNivel5.Text) _
    Or Not IsNumeric(txtNivel6.Text) Or Not IsNumeric(txtNivel7.Text) Or Not IsNumeric(txtNivel8.Text) Then
 vMensaje = vMensaje & vbCrLf & " - Valor(es) en el(los) Nivel(es) es inválido "
End If

If txtNivel1 = 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 1 no es válido "
If txtNivel1 = 0 And txtNivel2 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 1 no es válido "
If txtNivel2 = 0 And txtNivel3 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 2 no es válido "
If txtNivel3 = 0 And txtNivel4 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 3 no es válido "
If txtNivel4 = 0 And txtNivel5 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 4 no es válido "
If txtNivel5 = 0 And txtNivel6 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 5 no es válido "
If txtNivel6 = 0 And txtNivel7 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 6 no es válido "
If txtNivel7 = 0 And txtNivel8 > 0 Then vMensaje = vMensaje & vbCrLf & " - Nivel 7 no es válido "


If CInt(txtNivel1.Text) + CInt(txtNivel2.Text) + CInt(txtNivel3.Text) + CInt(txtNivel4.Text) + CInt(txtNivel5.Text) _
        + CInt(txtNivel6.Text) + CInt(txtNivel7.Text) + CInt(txtNivel8.Text) > 52 Then
   vMensaje = vMensaje & vbCrLf & " - Se excede el número maximo de caracteres soportado por el sistema"
End If

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre de la compañía no es válido "

If chkConsolida.Value = xtpChecked Then
   If cboConsolida_Conta.ListCount < 0 Then
        vMensaje = vMensaje & vbCrLf & " - No es posible referenciar a Niguna Contabilidad BASE para la consolidación!"
   End If
   
   If cboConsolida_Unidad.ListCount < 0 Then
        vMensaje = vMensaje & vbCrLf & " - No es posible referenciar a Niguna Unidad BASE para la consolidación!"
   End If
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function

Private Sub chkConsolida_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

vPaso = True

If chkConsolida.Value = xtpUnchecked Then
   
   cboConsolida_Conta.Clear
   cboConsolida_Unidad.Clear

Else
   strSQL = "exec spCntX_Consolida_Base_List " & vCodigo
   Call sbCbo_Llena_New(cboConsolida_Conta, strSQL, False, True)
   
   strSQL = "exec spCntX_Consolida_Unidades_List " & vCodigo
   Call sbCbo_Llena_New(cboConsolida_Unidad, strSQL, False, True)

End If


vPaso = False

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or Not IsNumeric(txtCodigo.Text) Then
   txtCodigo.Text = "0"
End If

If vScroll Then
    strSQL = "select Top 1 cod_Contabilidad,Nombre from Cntx_Contabilidades"

    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Contabilidad > " & txtCodigo.Text & " order by cod_Contabilidad asc"
    Else
       strSQL = strSQL & " where cod_Contabilidad < " & txtCodigo.Text & " order by cod_Contabilidad desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_CONTABILIDAD)
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
vModulo = 20
End Sub

Private Sub Form_Load()
 vModulo = 20
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = True
 
 cboRazon.Clear
 cboRazon.AddItem "Comercial"
 cboRazon.AddItem "Social"
 cboRazon.Text = "Comercial"
 
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()

vTipoBusca = "D"

vCodigo = 0
txtCodigo.Text = ""
txtNombre.Text = ""
txtEmail.Text = ""
txtContacto.Text = ""
txtMascara.Text = ""

txtNivel1.Text = 0
txtNivel2.Text = 0
txtNivel3.Text = 0
txtNivel4.Text = 0
txtNivel5.Text = 0
txtNivel6.Text = 0
txtNivel7.Text = 0
txtNivel8.Text = 0

txtTelCentral.Text = ""
txtTelFax.Text = ""
cboRazon.Text = "Comercial"

txtCodigo.Enabled = True

chkExpAreas.Value = vbChecked
chkExpAsientosGenerales.Value = vbChecked
chkExpCuentas.Value = vbChecked
chkExpDiferidos.Value = vbChecked
chkExpMantenimiento.Value = vbChecked
chkExpPlanFijos.Value = vbChecked
chkExpPlanRate.Value = vbChecked
chkExpPresupuesto.Value = vbChecked


chkFiltra_Bancos.Value = xtpUnchecked
chkFiltra_Contabilidad.Value = xtpUnchecked
chkFiltra_Inversiones.Value = xtpUnchecked
chkFiltra_Operaciones.Value = xtpUnchecked
chkFiltra_RRHH.Value = xtpUnchecked

chkConsolida.Value = xtpUnchecked

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    Case "CONSULTAR"
       If vTipoBusca = "D" Then
         gBusquedas.Columna = "nombre"
         gBusquedas.Orden = "nombre"
       Else
         gBusquedas.Columna = "COD_CONTABILIDAD"
         gBusquedas.Orden = "COD_CONTABILIDAD"
       End If
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select COD_CONTABILIDAD,nombre from cntX_contabilidades"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtNombre.SetFocus
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    Case "CERRAR"
      Unload Me
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_Contabilidad_Consulta " & lngCodigo
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  vEdita = True
  vCodigo = rs!COD_CONTABILIDAD
  'llenar datos en pantalla
  txtCodigo.Text = rs!COD_CONTABILIDAD
  txtNombre.Text = IIf(IsNull(rs!Nombre), "", rs!Nombre)
  txtCedulaJuridica.Text = IIf(IsNull(rs!cedula_juridica), "", rs!cedula_juridica)
  txtContacto.Text = IIf(IsNull(rs!contacto), "", rs!contacto)
  
  txtNivel1.Text = IIf(IsNull(rs!Nivel1), 0, rs!Nivel1)
  txtNivel2.Text = IIf(IsNull(rs!Nivel2), 0, rs!Nivel2)
  txtNivel3.Text = IIf(IsNull(rs!Nivel3), 0, rs!Nivel3)
  txtNivel4.Text = IIf(IsNull(rs!Nivel4), 0, rs!Nivel4)
  txtNivel5.Text = IIf(IsNull(rs!Nivel5), 0, rs!Nivel5)
  txtNivel6.Text = IIf(IsNull(rs!Nivel6), 0, rs!Nivel6)
  txtNivel7.Text = IIf(IsNull(rs!Nivel7), 0, rs!Nivel7)
  txtNivel8.Text = IIf(IsNull(rs!Nivel8), 0, rs!Nivel8)
  
  txtTelCentral.Text = IIf(IsNull(rs!tel_central), "", rs!tel_central)
  txtTelFax.Text = IIf(IsNull(rs!tel_fax), "", rs!tel_fax)
  
  txtEmail.Text = rs!Email & ""
  
  chkExpAreas.Value = rs!ExpAreas
  chkExpAsientosGenerales.Value = rs!ExpAsientos
  chkExpCuentas.Value = rs!ExpCuentas
  chkExpDiferidos.Value = rs!ExpDiferidos
  chkExpMantenimiento.Value = rs!ExpMantenimiento
  chkExpPlanFijos.Value = rs!ExpPlanFijo
  chkExpPlanRate.Value = rs!ExpPlanRate
  chkExpPresupuesto.Value = rs!ExpPresupuesto
  
    chkFiltra_Bancos.Value = rs!FILTRA_CTAS_BANCOS
    chkFiltra_Contabilidad.Value = rs!FILTRA_CTAS_CONTABILIDAD
    chkFiltra_Inversiones.Value = rs!FILTRA_CTAS_INVERSIONES
    chkFiltra_Operaciones.Value = rs!FILTRA_CTAS_OPERACIONES
    chkFiltra_RRHH.Value = rs!FILTRA_CTAS_RRHH
    
  Call sbCboAsignaDato(cboRazon, rs!RazonSocial_Desc, True, rs!RazonSocial)
  
  
  
    chkConsolida.Value = rs!Consolida_Ind
    Call chkConsolida_Click
  
  If rs!Consolida_Ind = 1 Then
        Call sbCboAsignaDato(cboConsolida_Conta, rs!ContaBase_Desc, True, rs!ContaBase_ID)
        Call sbCboAsignaDato(cboConsolida_Unidad, rs!Unidad_Desc, True, rs!Unidad_ID)
  End If
  
  On Error Resume Next
  txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")
  
Else
  MsgBox "No se encontró empresa contable verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxRazon(vRazon As String) As String

Select Case UCase(vRazon)
  Case "COMERCIAL"
    fxRazon = "C"
  Case "SOCIAL"
    fxRazon = "S"
  Case "C"
    fxRazon = "Comercial"
  Case "S"
    fxRazon = "Social"
End Select

End Function


Private Sub sbPredeterminados()
Dim iRespuesta As Integer, strSQL As String

On Error GoTo vError

iRespuesta = MsgBox("Desea crear Tipos de Cuentas y Tipos de Asientos Predeterminados por ProGrX: Contabilidad a Esta Contabilidad", vbYesNo)

If iRespuesta = vbYes Then
  
  Me.MousePointer = vbHourglass
  
    strSQL = "exec spCntX_Util_Contabilidad_Cfg_Predetermina " & txtCodigo.Text _
           & ",1,'" & glogon.Usuario & "','*xHM1tOk3n$'"
    Call ConectionExecute(strSQL)
       
  Me.MousePointer = vbDefault

Else
 
 MsgBox "TIENE QUE CREAR LOS TIPOS DE " _
    & "DIVISAS,CUENTAS Y ASIENTOS, Y LUEGO CONFIGURAR LOS OTROS INGRESOS " _
    & "Y OTROS GASTOS EN LA CONFIGURACION DEL ESTADO DE RESULTADO...", vbInformation

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxCatalogoLineas(pContabilidad As Long) As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Count(*) as Lineas from CntX_Cuentas where cod_contabilidad = " & pContabilidad
Call OpenRecordSet(rs, strSQL, 0)
    fxCatalogoLineas = rs!Lineas
rs.Close

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pConsolidaConta As String, pConsolidaUnidad As String

On Error GoTo vError

strSQL = "select isnull(count(*),0) as Total " _
       & " from CntX_Contabilidades" _
       & " Where nombre = '" & UCase(Trim(txtNombre.Text)) & "'"
Call OpenRecordSet(rs, strSQL)

If Not vEdita Then
  If rs!Total > 0 Then
    rs.Close
    MsgBox "El nombre de esta contabilidad ya se encuentra registrado verifique...", vbCritical
    Exit Sub
  End If
End If
rs.Close

If chkConsolida.Value = xtpChecked Then
        pConsolidaConta = cboConsolida_Conta.ItemData(cboConsolida_Conta.ListIndex)
        pConsolidaUnidad = "'" & cboConsolida_Unidad.ItemData(cboConsolida_Unidad.ListIndex) & "'"
        
Else
    pConsolidaConta = "0"
    pConsolidaUnidad = "''"
End If



If vEdita Then
  'Verificar si cambio cedula o codigo para actualización en cascada
  strSQL = "update CntX_Contabilidades set nombre = '" & Trim(txtNombre.Text) & "'" _
         & ", cedula_juridica = '" & txtCedulaJuridica.Text & "'" _
         & ", tel_fax = '" & txtTelFax.Text & "',tel_central = '" & txtTelCentral & "'" _
         & ", contacto = '" & txtContacto & "', Email = '" & txtEmail.Text & "'" _
         & ", razonsocial = '" & Mid(cboRazon.Text, 1, 1) & "'" _
         & ", hecho = '',revisado = ''" _
         & ", ExpAreas = " & chkExpAreas.Value & ",ExpAsientos = " & chkExpAsientosGenerales.Value _
         & ", ExpCuentas = " & chkExpCuentas.Value & ",ExpMantenimiento = " & chkExpMantenimiento.Value _
         & ", ExpDiferidos = " & chkExpDiferidos.Value & ",ExpPlanFijo = " & chkExpPlanFijos.Value _
         & ", ExpPlanRate = " & chkExpPlanRate.Value & ",expPresupuesto = " & chkExpPresupuesto.Value _
         & ", FILTRA_CTAS_BANCOS = " & chkFiltra_Bancos.Value & ", FILTRA_CTAS_CONTABILIDAD = " & chkFiltra_Contabilidad.Value _
         & ", FILTRA_CTAS_INVERSIONES = " & chkFiltra_Inversiones.Value & ", FILTRA_CTAS_OPERACIONES = " & chkFiltra_Operaciones.Value _
         & ", FILTRA_CTAS_RRHH = " & chkFiltra_RRHH.Value _
         & ", I_CONSOLIDADORA = " & chkConsolida.Value & ", CONSOLIDA_CONTA_BASE = " & pConsolidaConta & ", CONSOLIDA_UNIDAD_BASE = " & pConsolidaUnidad _
         & ", MODIFICA_FECHA = GETDATE(), MODIFICA_USUARIO = '" & glogon.Usuario & "'"
         
  'Si el catalogo ya tiene cuentas registradas no permite actualizar niveles
  If fxCatalogoLineas(vCodigo) = 0 Then
    strSQL = strSQL & ",Nivel1 = " & txtNivel1.Text & ",Nivel2 = " & txtNivel2.Text & ",Nivel3 = " & txtNivel3.Text _
           & ",Nivel4 = " & txtNivel4.Text & ",Nivel5 = " & txtNivel5.Text & ", Nivel6 = " & txtNivel6.Text _
           & ",Nivel7 = " & txtNivel7.Text & ",Nivel8 = " & txtNivel8.Text
  End If
  strSQL = strSQL & " where COD_CONTABILIDAD = " & vCodigo
        
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Modifica", "Contabilidad : " & vCodigo)

Else
   
   strSQL = "select isnull(max(COD_CONTABILIDAD),0) + 1 as ultimo from CntX_Contabilidades"
   Call OpenRecordSet(rs, strSQL, 0)
     txtCodigo = rs!ultimo
     vCodigo = txtCodigo
   rs.Close
   
   
   strSQL = "insert into CntX_Contabilidades(cod_contabilidad,nombre,cedula_juridica,tel_fax,tel_central,contacto, email" _
          & ", razonsocial,nivel1,nivel2,nivel3,nivel4,nivel5,nivel6, nivel7, nivel8, hecho,revisado" _
          & ", ExpAreas,ExpCuentas,ExpAsientos,ExpMantenimiento,ExpDiferidos" _
          & ", ExpPlanFijo,ExpPlanRate,ExpPresupuesto, FILTRA_CTAS_BANCOS, FILTRA_CTAS_CONTABILIDAD" _
          & ", FILTRA_CTAS_INVERSIONES, FILTRA_CTAS_OPERACIONES, FILTRA_CTAS_RRHH" _
          & ", I_CONSOLIDADORA, CONSOLIDA_CONTA_BASE, CONSOLIDA_UNIDAD_BASE, registro_Fecha, Registro_Usuario)" _
          & " values(" & vCodigo & ",'" & Trim(txtNombre.Text) & "','" & Trim(txtCedulaJuridica.Text) _
          & "', '" & txtTelFax & "','" & txtTelCentral & "','" & Trim(txtContacto.Text) & "','" & RTrim(txtEmail.Text) _
          & "', '" & fxRazon(cboRazon.Text) & "'," & txtNivel1.Text & "," & txtNivel2.Text _
          & ", " & txtNivel3.Text & "," & txtNivel4.Text & "," & txtNivel5.Text & "," & txtNivel6.Text & "," & txtNivel7.Text & "," & txtNivel8.Text _
          & ", '',''," & chkExpAreas.Value & "," & chkExpCuentas.Value & "," & chkExpAsientosGenerales.Value _
          & ", " & chkExpMantenimiento.Value & "," & chkExpDiferidos.Value _
          & ", " & chkExpPlanFijos.Value & "," & chkExpPlanRate.Value & "," & chkExpPresupuesto.Value _
          & ", " & chkFiltra_Bancos.Value & ", " & chkFiltra_Contabilidad.Value _
          & ", " & chkFiltra_Inversiones.Value & ", " & chkFiltra_Operaciones.Value & ", " & chkFiltra_RRHH.Value _
          & ", " & chkConsolida.Value & ", " & pConsolidaConta & ", " & pConsolidaUnidad _
          & ", getdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL, 0)
    
   Call sbPredeterminados
    
   Call Bitacora("Registra", "Contabilidad : " & txtCodigo)
    
   MsgBox "RECUERDE QUE TIENE QUE DIFINIR LOS PERIODOS Y EL CIERRE FISCAL PARA ESTA CONTABILIDAD...", vbInformation
    
   txtCodigo.Enabled = True
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CntX_Contabilidades where COD_CONTABILIDAD = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Elimina", "Empresa Contable : " & vCodigo)

  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbUsuarios_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbClassCall("Contabilidad", 0, "frmCntX_ContabilidadesUsuarios")
End Sub

Private Sub txtCedulaJuridica_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtContacto.SetFocus
End Sub


Private Sub txtCodigo_GotFocus()
 vTipoBusca = "C"
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtContacto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtTelCentral.SetFocus
End Sub

Private Sub txtNivel1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel2.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub

Private Sub txtNivel2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel3.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")
End Sub

Private Sub txtNivel3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel4.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub

Private Sub txtNivel4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel5.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub

Private Sub txtNivel5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel6.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub


Private Sub txtNivel6_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel7.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub


Private Sub txtNivel7_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel8.SetFocus
txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub


Private Sub txtNivel8_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

txtMascara.Text = fxCntX_CuentaMascara(txtNivel1.Text, txtNivel2.Text, txtNivel3.Text, txtNivel4.Text, txtNivel5.Text, txtNivel6.Text, txtNivel7.Text, txtNivel8.Text, "0")

End Sub


Private Sub txtNombre_GotFocus()
 vTipoBusca = "D"
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtCedulaJuridica.SetFocus
End Sub


Private Sub txtTelCentral_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtTelFax.SetFocus
End Sub

Private Sub txtTelFax_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cboRazon.SetFocus
End Sub


Private Sub cboRazon_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNivel1.SetFocus
End Sub

