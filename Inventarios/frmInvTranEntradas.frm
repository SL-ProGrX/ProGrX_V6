VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmInvTranEntradas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario: ENTRADAS"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12015
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   12015
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   3555
      NewRow1         =   0   'False
      Child2          =   "tlbProcesos"
      MinHeight2      =   330
      Width2          =   3645
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbProcesos 
         Height          =   330
         Left            =   3750
         TabIndex        =   2
         Top             =   30
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   582
         ButtonWidth     =   1958
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Plantillas"
               Key             =   "Plantilla"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Autoriza"
               Key             =   "Autoriza"
               Object.ToolTipText     =   "Autorizacion o Rechazo"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Procesar"
               Key             =   "Procesar"
               Object.ToolTipText     =   "Ejecuta la Transaccion"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
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
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta "
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repListadoGeneral"
                     Text            =   "Listado General"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTranEntradas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTranEntradas.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTranEntradas.frx":0BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTranEntradas.frx":0F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvTranEntradas.frx":1228
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.CheckBox chkPlantilla 
      Height          =   252
      Left            =   6360
      TabIndex        =   3
      Top             =   600
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Plantilla?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   5640
      TabIndex        =   4
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   312
      Left            =   9480
      TabIndex        =   5
      Top             =   600
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
      Height          =   432
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   312
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbDetalle 
      Height          =   1212
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   11772
      _Version        =   1441793
      _ExtentX        =   20764
      _ExtentY        =   2138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   9480
         TabIndex        =   12
         Top             =   240
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   552
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   9612
         _Version        =   1441793
         _ExtentX        =   16954
         _ExtentY        =   974
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
      Begin XtremeSuiteControls.ComboBox cboCausa 
         Height          =   312
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Causa: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   2
         Left            =   7920
         TabIndex        =   15
         Top             =   240
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "No. Documento: "
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
      End
   End
   Begin XtremeSuiteControls.GroupBox gbResume 
      Height          =   972
      Left            =   0
      TabIndex        =   18
      Top             =   5760
      Width           =   11892
      _Version        =   1441793
      _ExtentX        =   20976
      _ExtentY        =   1714
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtSubTotal 
         Height          =   312
         Left            =   9840
         TabIndex        =   19
         Top             =   480
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblCantidad 
         Height          =   492
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cantidad 0"
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
         Height          =   492
         Left            =   8880
         TabIndex        =   21
         Top             =   360
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Total"
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
      Begin XtremeSuiteControls.Label lblLineas 
         Height          =   492
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4890
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Lineas 0"
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
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   6780
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Solicitado Por"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Solicitado Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Resuelto Por"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Resuelto Fecha"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Procesado Por"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Procesado Fecha"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3252
      Left            =   0
      TabIndex        =   24
      Top             =   2400
      Width           =   11892
      _Version        =   524288
      _ExtentX        =   20976
      _ExtentY        =   5736
      _StockProps     =   64
      ArrowsExitEditMode=   -1  'True
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
      MaxCols         =   487
      ScrollBars      =   2
      SpreadDesigner  =   "frmInvTranEntradas.frx":1B02
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   3
      Left            =   7920
      TabIndex        =   9
      Top             =   600
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha: "
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Boleta: "
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
End
Attribute VB_Name = "frmInvTranEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vTipo As String
Dim vMascara As String, vScroll As Boolean

Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 Boleta from pv_invTransac" _
           & " where Tipo = '" & vTipo & "'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and boleta > '" & Format(txtCodigo, vMascara) & "' order by Boleta asc"
    Else
       strSQL = strSQL & " and boleta < '" & Format(txtCodigo, vMascara) & "' order by Boleta desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!Boleta)
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
vModulo = 32
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 32
 vTipo = "E"
 vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 
 vMascara = "0000000000"
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = ""
txtCodigo = ""

txtDocumento = ""
txtFecha = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
dtpFecha.Value = fxFechaServidor
dtpFecha.Visible = fxInvCambiaFecha(glogon.Usuario, vTipo)

txtNotas = ""

txtEstado = ""
txtEstado.Tag = "S"

Call sbInvESCombo(vTipo, cboCausa)

vGrid.MaxRows = 1
vGrid.MaxCols = 6
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0

txtCodigo.Enabled = True

With StatusBarX.Panels
  .Item(1).Text = ""
  .Item(2).Text = ""
  .Item(3).Text = ""
  .Item(4).Text = ""
  .Item(5).Text = ""
End With

tlbProcesos.Buttons.Item(2).Enabled = False
tlbProcesos.Buttons.Item(3).Enabled = False

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      cboCausa.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNotas.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
        Call txtCodigo_KeyDown(vbKeyF4, 1)
        
    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select X.*,rtrim(C.descripcion) as Causa" _
       & " from PV_INVTRANSAC X inner join pv_entrada_salida C on X.cod_entsal = C.cod_entsal" _
       & " where X.boleta = '" & Format(xCodigo, vMascara) & "' and X.tipo = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!Boleta
  txtCodigo = rs!Boleta
  
  txtDocumento = rs!Documento
  
  Call sbCboAsignaDato(cboCausa, rs!Causa, True, Trim(rs!COD_ENTSAL))
  
 
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  dtpFecha.Value = rs!fecha
  
  txtNotas = rs!Notas & ""
    
  tlbProcesos.Buttons.Item(2).Enabled = False
  tlbProcesos.Buttons.Item(3).Enabled = False
  
  chkPlantilla.Value = rs!plantilla
  
  Select Case rs!Estado
    Case "S" 'Solicitada
        txtEstado = "Solicitada"
        tlbProcesos.Buttons.Item(2).Enabled = True
    Case "A" 'Autorizada
        txtEstado = "Autorizada"
        tlbProcesos.Buttons.Item(3).Enabled = True
    Case "R" 'Rechazada
        txtEstado = "Rechazada"
    Case "P" 'Procesada
        txtEstado = "Procesada"
  End Select
  
  txtEstado.Tag = rs!Estado
    
  With StatusBarX.Panels
    .Item(1) = rs!genera_user
    .Item(2) = rs!genera_fecha
    .Item(3) = rs!Autoriza_user & ""
    .Item(4) = rs!Autoriza_Fecha & ""
    .Item(5) = rs!Procesa_user & ""
    .Item(6) = rs!Procesa_Fecha & ""
  End With
    

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,B.cod_bodega,B.descripcion as Bodega,D.precio,(D.cantidad * D.precio) as Total" _
         & ",isnull(D.despacho,0) as Despacho" _
         & " from PV_INVTRADET D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " inner join PV_Bodegas B on D.cod_bodega = B.cod_bodega" _
         & " where D.boleta = '" & rs!Boleta & "' and D.tipo = '" & rs!Tipo & "'"
         
  Call sbCargaGridLocal(vGrid, 6, strSQL)
  
  Call sbCalculaTotales
    
  vGrid.Enabled = True
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Public Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
End If
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
        Select Case i
          Case 1
              vGrid.col = i
              vGrid.Text = CStr(rs!Cod_Producto)
         
          Case 2
              vGrid.col = i
              vGrid.Text = CStr(rs!Descripcion)
             
          Case 3
              vGrid.col = i
              vGrid.Text = CStr(rs!Cantidad)
              vGrid.TextTip = TextTipFixed
              vGrid.CellNote = "Unidades Despachadas : " & rs!despacho & vbCrLf & " Unidades Pendientes : " _
                             & (rs!Cantidad - rs!despacho)
          
          Case 4
              vGrid.col = i
              vGrid.Text = CStr(rs!cod_bodega)
              vGrid.TextTip = TextTipFixed
              vGrid.CellNote = rs!Bodega
          
          Case 5
              vGrid.col = i
              vGrid.Text = CStr(rs!Precio)
          
          Case 6
              vGrid.col = i
              vGrid.Text = CStr(rs!Total)
          
        End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

On Error GoTo vError

vMensaje = fxInvVerificaLineaDetalle(vGrid, 3, "S", 1, 4)

If Not fxInvPeriodos(dtpFecha.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."

vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, i As Integer
Dim rs As New ADODB.Recordset
Dim vFecha As Date, curCantidad As Currency

On Error GoTo vError


If txtEstado.Tag <> "S" Then
  MsgBox "Esta Transaccion no esta solicitada, No se puede Modificar...", vbExclamation
  Exit Sub
End If

If dtpFecha.Visible Then
  vFecha = dtpFecha.Value
Else
  vFecha = fxFechaServidor
End If


If vEdita Then
    strSQL = "update pv_InvTranSac set cod_entsal = '" & cboCausa.ItemData(cboCausa.ListIndex) _
           & "',fecha = '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',documento = '" _
           & txtDocumento & "', notas = '" & txtNotas & "',plantilla = " & chkPlantilla.Value _
           & ", Total = " & CCur(txtSubTotal.Text) _
           & " where Boleta = '" & vCodigo & "' and Tipo = '" & vTipo & "'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Transaccion Inv.Tipo: " & vTipo & ", Cod: " & vCodigo)

Else
    
    strSQL = "select isnull(max(Boleta),0)+1 as Ultimo from pv_InvTranSac where Tipo = '" & vTipo & "'"
    Call OpenRecordSet(rs, strSQL)
      vCodigo = Format(rs!ultimo, vMascara)
    rs.Close
    txtCodigo = vCodigo
    
    strSQL = "insert pv_InvTranSac(Boleta,Tipo,cod_entsal,genera_fecha,documento,notas" _
           & ",genera_user,estado,plantilla,fecha,fecha_sistema,total)" _
           & " values('" & vCodigo & "','" & vTipo & "','" & cboCausa.ItemData(cboCausa.ListIndex) _
           & "',dbo.MyGetdate(),'" & txtDocumento & "','" & txtNotas _
           & "','" & glogon.Usuario & "','S'," & chkPlantilla.Value & ",'" _
           & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',dbo.MyGetdate()," & CCur(txtSubTotal.Text) & ")"
    
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", "Transaccion Inv.Tipo: " & vTipo & ", Cod: " & vCodigo)

End If

txtCodigo.Enabled = True

'Guardar Detalle de la Transaccion
strSQL = "delete pv_InvTraDet where tipo = '" & vTipo & "' and boleta = '" & vCodigo & "'"

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    strSQL = strSQL & Space(10) & "insert pv_InvTraDet(linea,Boleta,tipo,cod_producto,cod_bodega,cantidad,despacho,precio)" _
           & " values(" & i & ",'" & vCodigo & "','" & vTipo & "','" & vGrid.Text & "','"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "'," & curCantidad & ",0,"
    vGrid.col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ")"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
      strSQL = ""
    End If
  
  End If
Next i

'Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


Call sbConsulta(vCodigo)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
  strSQL = "delete pv_InvTranSac where tipo = '" & vTipo & "' and Boleta = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Transac.Inv.Tipo (" & vTipo & ") Cod." & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)


InvTransacRep.Tipo = vTipo
InvTransacRep.Boleta = vCodigo
InvTransacRep.Reporte = ""

Select Case UCase(ButtonMenu.Key)
  Case "REPBOLETA"
     frmInvTransacReporteOrden.Show vbModal
  
  Case "REPLISTADOGENERAL"
     Call sbFormsCall("frmInvTransacReportes", 1)

End Select

End Sub


Private Sub tlbProcesos_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError


Select Case Button.Key
  Case "Plantilla"
        gBusquedas.Columna = "Boleta"
        gBusquedas.Orden = "Boleta"
        gBusquedas.Mascara = vMascara
        gBusquedas.Consulta = "select Boleta,genera_user,genera_fecha,documento,notas" _
                  & " from pv_InvTransac"
        gBusquedas.Filtro = " and plantilla = 1 and tipo = '" & vTipo & "'"
        gBusquedas.Resultado = ""
        gBusquedas.Resultado2 = ""
        frmBusquedas.Show vbModal
        
        If gBusquedas.Resultado <> "" Then
          Call sbLimpiaPantalla
          Call sbConsulta(gBusquedas.Resultado)
          txtCodigo = ""
          chkPlantilla.Value = vbUnchecked
          vEdita = False
          txtCodigo.Enabled = False
          cboCausa.SetFocus
          Call sbToolBar(tlb, "edicion")
          
          txtEstado = ""
          txtEstado.Tag = "S"
          tlbProcesos.Buttons.Item(2).Enabled = False
          tlbProcesos.Buttons.Item(3).Enabled = False
       End If
  
  Case "Autoriza"
       gInvTran.Boleta = vCodigo
       gInvTran.Tipo = vTipo
       gInvTran.fecha = dtpFecha.Value
       gInvTran.Causa = cboCausa.Text
       frmInvTransacAutoriza.Show vbModal
       Call sbConsulta(vCodigo)
  
  Case "Procesar"
       gInvTran.Boleta = vCodigo
       gInvTran.Tipo = vTipo
       gInvTran.fecha = dtpFecha.Value
       gInvTran.Causa = cboCausa.Text
       frmInvTransacProcesa.Show vbModal
       Call sbConsulta(vCodigo)
  
End Select

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus

If KeyCode = vbKeyF4 Then
  gInvTran.Boleta = vCodigo
  gInvTran.Tipo = vTipo
  frmInvTransacQry.Show vbModal
  txtCodigo = gInvTran.Boleta
  If txtCodigo <> "" Then Call sbConsulta(gInvTran.Boleta)
  chkPlantilla.Enabled = True
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub sbCalculaTotales()
Dim curSubTotal As Currency
Dim curTmpPrecio As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long
Dim iLineas As Integer, curCantidad As Currency


curSubTotal = 0

iLineas = 0
curCantidad = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)

    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curCantidad = curCantidad + curTmpCant
    iLineas = iLineas + 1
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")

lblLineas.Caption = "Líneas: " & iLineas
lblCantidad.Caption = "Cantidad: " & Format(curCantidad, "Standard")

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,costo_regular,impuesto_ventas from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!Cod_Producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 5
  vGrid.Text = CStr(rs!costo_regular)
  vGrid.col = 3
  vGrid.Text = 1
End If
rs.Close


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consulta Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_salidas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 4
   vGrid.Text = gBusquedas.Resultado
   vGrid.CellNote = gBusquedas.Resultado2
   vGrid.TextTip = TextTipFixed
End If

If vGrid.ActiveCol = 4 And KeyCode = vbKeyReturn Then
   vGrid.col = 4
   vGrid.CellNote = fxSIFCCodigos("D", vGrid.Text, "bodegas")
   vGrid.TextTip = TextTipFixed
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.DeleteRows vGrid.ActiveRow, 1
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  Call sbCalculaTotales
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub


Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim curCantidad As Currency, curPrecio As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    vGrid.Text = (curPrecio * curCantidad)
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:

End Sub






