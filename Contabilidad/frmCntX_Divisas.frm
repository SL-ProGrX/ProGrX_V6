VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCntX_Divisas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definición de Divisas"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerInicio 
      Interval        =   10
      Left            =   9600
      Top             =   240
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5172
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9972
      _Version        =   1441792
      _ExtentX        =   17590
      _ExtentY        =   9123
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
      Item(0).Caption =   "General"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "gbDiferencia"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "cboSimbolo"
      Item(0).Control(3)=   "Label3(2)"
      Item(0).Control(4)=   "chkDivisa"
      Item(0).Control(5)=   "txtNotas"
      Item(0).Control(6)=   "gbTipoCambioDefaul"
      Item(1).Caption =   "Diferenciales"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Item(2).Caption =   "TIpos de Cambios"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "lswTC"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4812
         Left            =   -70000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   9972
         _Version        =   1441792
         _ExtentX        =   17590
         _ExtentY        =   8488
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswTC 
         Height          =   4812
         Left            =   -70000
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   9972
         _Version        =   1441792
         _ExtentX        =   17590
         _ExtentY        =   8488
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
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox gbTipoCambioDefaul 
         Height          =   1212
         Left            =   120
         TabIndex        =   23
         Top             =   3960
         Width           =   9732
         _Version        =   1441792
         _ExtentX        =   17166
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Tipo de Cambio por Omisión"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtTC_Venta 
            Height          =   312
            Left            =   4200
            TabIndex        =   26
            Top             =   360
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtTC_Compra 
            Height          =   312
            Left            =   4200
            TabIndex        =   27
            Top             =   720
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   2
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Cambio, Compra:"
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
            Left            =   1680
            TabIndex        =   25
            Top             =   720
            Width           =   2172
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo Cambio, Venta:"
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
            Left            =   1680
            TabIndex        =   24
            Top             =   360
            Width           =   2172
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDivisa 
         Height          =   252
         Left            =   3240
         TabIndex        =   17
         Top             =   480
         Width           =   2292
         _Version        =   1441792
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Divisa Funcional ?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox gbDiferencia 
         Height          =   2052
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   9732
         _Version        =   1441792
         _ExtentX        =   17166
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Registro por Omisión del diferencial cambiario: "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtDC_Cuenta_GST_Desc 
            Height          =   312
            Left            =   3720
            TabIndex        =   4
            Top             =   840
            Width           =   5412
            _Version        =   1441792
            _ExtentX        =   9546
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDC_Cuenta_ING 
            Height          =   312
            Left            =   1680
            TabIndex        =   5
            Top             =   480
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDC_Cuenta_ING_Desc 
            Height          =   312
            Left            =   3720
            TabIndex        =   6
            Top             =   480
            Width           =   5412
            _Version        =   1441792
            _ExtentX        =   9546
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtDC_Cuenta_GST 
            Height          =   312
            Left            =   1680
            TabIndex        =   7
            Top             =   840
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.ComboBox cboUnidad 
            Height          =   312
            Left            =   1680
            TabIndex        =   8
            Top             =   1200
            Width           =   4572
            _Version        =   1441792
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboCentro 
            Height          =   312
            Left            =   1680
            TabIndex        =   9
            Top             =   1560
            Width           =   4572
            _Version        =   1441792
            _ExtentX        =   8070
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad"
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
            Left            =   360
            TabIndex        =   22
            Top             =   1200
            Width           =   1572
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Costo"
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
            Left            =   360
            TabIndex        =   21
            Top             =   1560
            Width           =   1572
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.: Ingresos"
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
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Width           =   1572
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cta.: Gastos"
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
            Left            =   360
            TabIndex        =   19
            Top             =   840
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.ComboBox cboSimbolo 
         Height          =   312
         Left            =   1800
         TabIndex        =   15
         Top             =   480
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   7572
         _Version        =   1441792
         _ExtentX        =   13356
         _ExtentY        =   1397
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
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Simbolo"
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
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   852
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   852
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9600
      TabIndex        =   11
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   960
      TabIndex        =   12
      Top             =   600
      Width           =   972
      _Version        =   1441792
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   1920
      TabIndex        =   13
      Top             =   600
      Width           =   7572
      _Version        =   1441792
      _ExtentX        =   13356
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
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
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   852
   End
End
Attribute VB_Name = "frmCntX_Divisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vTipoBusca As String
Dim vScroll As Boolean


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre de la divisa no es válido "
If txtDC_Cuenta_ING = "" Then
  'vMensaje = vMensaje & vbCrLf & " - La cuenta contable no es válida "
  MsgBox "Advertencia...: Se debe de configurar la cuenta contable antes de realizar ajustes de diferencial cambiario.!", vbExclamation
End If

If Not IsNumeric(txtTC_Venta.Text) Then vMensaje = vMensaje & vbCrLf & " - Tipo de Cambio Inicial de Venta no es válido "
If Not IsNumeric(txtTC_Compra.Text) Then vMensaje = vMensaje & vbCrLf & " - Tipo de Cambio Inicial de Venta no es válido "


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function


Private Sub sbDiferenciales_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

With lsw.ColumnHeaders
    .Clear
    .Add , , "Corte", 1400, vbCenter
    .Add , , "TC Compra", 1800, vbRightJustify
    .Add , , "TC Venta", 1800, vbRightJustify
    .Add , , "Año", 1000, vbCenter
    .Add , , "Mes", 1000, vbCenter
    .Add , , "Usuario", 2100
End With

lsw.ListItems.Clear

strSQL = "select * from CntX_Divisas_historial where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and cod_divisa = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Format(rs!fecha, "yyyy/mm/dd"))
     itmX.SubItems(1) = Format(rs!tc_compra, "Standard")
     itmX.SubItems(2) = Format(rs!tc_venta, "Standard")
     itmX.SubItems(3) = rs!Anio
     itmX.SubItems(4) = rs!Mes
     itmX.SubItems(5) = rs!Usuario
 rs.MoveNext
Loop
rs.Close

End Sub


Private Sub sbTiposCambio_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

With lswTC.ColumnHeaders
    .Clear
    .Add , , "Inicio", 1400, vbCenter
    .Add , , "Corte", 1400, vbCenter
    .Add , , "TC Compra", 1400, vbRightJustify
    .Add , , "TC Venta", 1400, vbRightJustify
    .Add , , "Usuario", 2000
    .Add , , "Fecha", 2000
End With

lswTC.ListItems.Clear

strSQL = "select Top 100 * from CNTX_DIVISAS_TIPO_CAMBIO" _
       & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and cod_divisa = '" & txtCodigo.Text & "' order by Corte Desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswTC.ListItems.Add(, , Format(rs!Inicio, "yyyy/mm/dd"))
     itmX.SubItems(1) = Format(rs!Corte, "yyyy/mm/dd")
     itmX.SubItems(2) = Format(rs!tc_compra, "Standard")
     itmX.SubItems(3) = Format(rs!tc_venta, "Standard")
     itmX.SubItems(4) = rs!Usuario & ""
     itmX.SubItems(5) = rs!fecha & ""
 rs.MoveNext
Loop
rs.Close

End Sub



Private Sub chkDivisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_divisa,descripcion from Cntx_Divisas" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_divisa > '" & txtCodigo.Text & "' order by cod_divisa asc"
    Else
       strSQL = strSQL & " and cod_divisa < '" & txtCodigo.Text & "' order by cod_divisa desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_DIVISA)
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

Private Sub Form_Load()

 vEdita = True
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 

 cboSimbolo.Clear
 cboSimbolo.AddItem ""
 cboSimbolo.AddItem "¢"
 cboSimbolo.AddItem "$"
 cboSimbolo.AddItem "€"
 cboSimbolo.AddItem "£"
 cboSimbolo.AddItem "¥"
 cboSimbolo.AddItem "Fr"
 cboSimbolo.AddItem "Kr"
 cboSimbolo.AddItem "W"
 cboSimbolo.AddItem ""
 cboSimbolo.Text = ""
 
 
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()

vTipoBusca = "D"

vCodigo = ""

txtCodigo = ""
txtDescripcion = ""
txtNotas = ""
txtTC_Compra.Text = "1"
txtTC_Venta.Text = "1"

cboSimbolo.Text = ""

txtDC_Cuenta_ING = ""
txtDC_Cuenta_ING_Desc = ""

txtDC_Cuenta_GST.Text = ""
txtDC_Cuenta_GST_Desc.Text = ""

txtTC_Compra.Locked = True
txtTC_Venta.Locked = True

chkDivisa.Value = vbUnchecked

tcMain.Item(0).Selected = True

End Sub



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 1
        Call sbDiferenciales_Consulta
    Case 2
        Call sbTiposCambio_Consulta
End Select

End Sub

Private Sub TimerInicio_Timer()

Dim strSQL As String

TimerInicio.Interval = 0
TimerInicio.Enabled = False

On Error GoTo vError

strSQL = "select rtrim(cod_unidad) as 'IdX', rtrim(descripcion) as ItmX" _
       & " from CntX_Unidades where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)

strSQL = "select COD_CENTRO_COSTO AS 'IdX', RTRIM(DESCRIPCION) AS 'ItmX'" _
       & " From CNTX_CENTRO_COSTOS" _
       & " Where Activo = 1 And COD_CONTABILIDAD = 1"
Call sbCbo_Llena_New(cboCentro, strSQL, False, True)
Call sbCboAsignaDato(cboCentro, "", True, "")

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
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
       If vTipoBusca = "D" Then
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       Else
         gBusquedas.Columna = "cod_divisa"
         gBusquedas.Orden = "cod_divisa"
       End If
       gBusquedas.Filtro = " and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
       gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
       frmBusquedas.Show vbModal
       txtCodigo = IIf((gBusquedas.Resultado = ""), "", gBusquedas.Resultado)
       Call sbConsulta(txtCodigo)
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    Case "CERRAR"
      UnLoad Me
End Select

End Sub

Private Sub sbConsulta(pDivisa As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select M.*, isnull(I.Cod_Cuenta_Mask,'') as 'CtaIng', isnull(G.Cod_Cuenta_Mask,'') as 'CtaGst'" _
       & ", isnull(I.descripcion,'') as 'CtaIng_Desc', isnull(G.Descripcion,'') as 'CtaGst_Desc'" _
       & ", isnull(U.Descripcion,'') as 'Unidad_Desc', isnull(Cc.Descripcion,'') as 'Centro_Desc'" _
       & " from CntX_Divisas M" _
       & " left join CntX_Cuentas I on M.COD_CONTABILIDAD = I.COD_CONTABILIDAD and M.cod_cuenta = I.cod_cuenta" _
       & " left join CntX_Cuentas G on M.COD_CONTABILIDAD = G.COD_CONTABILIDAD and M.cod_cuenta_Gasto = G.cod_cuenta" _
       & " left join CntX_Unidades U on M.COD_CONTABILIDAD = U.COD_CONTABILIDAD and M.cod_Unidad = U.cod_Unidad" _
       & " left join CntX_Centro_Costos Cc on M.COD_CONTABILIDAD = Cc.COD_CONTABILIDAD and M.cod_Centro_Costo = Cc.Cod_Centro_Costo" _
       & " where M.cod_divisa = '" & pDivisa & "' and M.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!COD_DIVISA
  
  'llenar datos en pantalla
  txtCodigo.Text = rs!COD_DIVISA
  txtDescripcion.Text = rs!Descripcion
      
  txtDescripcion.SetFocus
  
  txtNotas.Text = rs!observacion
  
  chkDivisa.Value = rs!divisa_local
  
  txtTC_Compra.Text = Format(rs!tc_compra, "Standard")
  txtTC_Venta.Text = Format(rs!tc_venta, "Standard")
  
  txtDC_Cuenta_ING.Text = rs!CtaIng
  txtDC_Cuenta_ING_Desc.Text = rs!CtaIng_Desc & ""
  
  txtDC_Cuenta_GST.Text = rs!CtaGst
  txtDC_Cuenta_GST_Desc.Text = rs!CtaGst_Desc & ""
  
  Call sbCboAsignaDato(cboUnidad, rs!Unidad_Desc, True, rs!cod_unidad & "")
  Call sbCboAsignaDato(cboCentro, rs!Centro_Desc, True, rs!cod_centro_costo & "")
  Call sbCboAsignaDato(cboSimbolo, Trim(rs!CURRENCY_SIM & ""), True, Trim(rs!CURRENCY_SIM & ""))
  
  
  
Else
  MsgBox "No se encontró registro de la divisa verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDecimal
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  
  strSQL = "update CntX_Divisas set descripcion = '" & Trim(txtDescripcion) _
         & "',observacion = '" & Trim(txtNotas.Text) _
         & "',cod_cuenta = '" & fxCntX_CuentaFormato(False, Trim(txtDC_Cuenta_ING.Text)) _
         & "',cod_cuenta_gasto = '" & fxCntX_CuentaFormato(False, Trim(txtDC_Cuenta_GST.Text)) _
         & "',divisa_local = " & chkDivisa.Value _
         & ", CURRENCY_SIM = '" & cboSimbolo.Text & "'" _
         & ", cod_Unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
         & ", cod_Centro_Costo = '" & cboCentro.ItemData(cboCentro.ListIndex) & "'" _
         & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " and cod_divisa = '" _
         & vCodigo & "'"
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Modifica", "Divisa : " & vCodigo & " Conta." & gCntX_Parametros.CodigoConta)

Else
   
   strSQL = "insert CntX_Divisas(COD_CONTABILIDAD,cod_divisa,descripcion,observacion,divisa_local,tc_compra" _
          & ",tc_venta,consecutivo,cod_cuenta,cod_cuenta_gasto,CURRENCY_SIM, cod_Unidad, Cod_Centro_Costo)" _
          & "  values(" & gCntX_Parametros.CodigoConta & ",'" & Trim(txtCodigo.Text) & "','" & Trim(txtDescripcion.Text) _
          & "','" & Trim(txtNotas.Text) & "'," & chkDivisa.Value _
          & "," & CCur(txtTC_Compra.Text) & "," & CCur(txtTC_Venta.Text) & ",0,'" _
          & fxCntX_CuentaFormato(False, Trim(txtDC_Cuenta_ING.Text)) & "','" _
          & fxCntX_CuentaFormato(False, Trim(txtDC_Cuenta_GST.Text)) & "','" _
          & cboSimbolo.Text & "','" _
          & cboUnidad.ItemData(cboUnidad.ListIndex) & "','" _
          & cboCentro.ItemData(cboCentro.ListIndex) & "')"
   Call ConectionExecute(strSQL, 0)
    
   vCodigo = txtCodigo
    
   Call Bitacora("Registra", "Divisa : " & txtCodigo & " Conta." & gCntX_Parametros.CodigoConta)
    
End If

If chkDivisa.Value = vbChecked Then
  strSQL = "update CntX_Divisas set divisa_local = 0 where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
         & " and cod_divisa not in('" & vCodigo & "')"
  Call ConectionExecute(strSQL, 0)
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
  strSQL = "delete CntX_Divisas where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
         & " and cod_divisa = '" & vCodigo & "'"
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Elimina", "Divisa : " & vCodigo & " Conta." & gCntX_Parametros.CodigoConta)
  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_GotFocus()
 vTipoBusca = "C"
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_divisa"
   gBusquedas.Orden = "cod_divisa"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
   gBusquedas.Filtro = " and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   frmBusquedas.Show vbModal
   txtCodigo = gBusquedas.Resultado
   txtDescripcion = gBusquedas.Resultado2
   Call sbConsulta(txtCodigo)
End If
End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDC_Cuenta_GST_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtDC_Cuenta_GST.Text, 0)
  
  txtDC_Cuenta_GST_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_GST.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:
End Sub

Private Sub txtDC_Cuenta_ING_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDC_Cuenta_ING_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDC_Cuenta_ING_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_ING.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtDC_Cuenta_ING_Desc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDC_Cuenta_GST.SetFocus

If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDC_Cuenta_ING_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_ING.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub


Private Sub txtDC_Cuenta_GST_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDC_Cuenta_GST_Desc.SetFocus
    
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDC_Cuenta_GST_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_GST.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub

Private Sub txtDC_Cuenta_GST_Desc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtDC_Cuenta_GST_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_GST.Text = fxCntX_CuentaFormato(True, gCuenta)
End If

End Sub



Private Sub txtDC_Cuenta_ING_LostFocus()
On Error GoTo vError
  
  gCuenta = fxCntX_CuentaFormato(False, txtDC_Cuenta_ING, 0)
  
  txtDC_Cuenta_ING_Desc.Text = fxCntX_Cuenta("D", gCuenta)
  txtDC_Cuenta_ING.Text = fxCntX_CuentaFormato(True, gCuenta)

Exit Sub
vError:
End Sub

Private Sub txtDescripcion_GotFocus()
 vTipoBusca = "D"
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkDivisa.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
   gBusquedas.Filtro = " and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   frmBusquedas.Show vbModal
   txtCodigo = gBusquedas.Resultado
   txtDescripcion = gBusquedas.Resultado2
   Call sbConsulta(txtCodigo)
End If
End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDC_Cuenta_ING.SetFocus
End Sub
