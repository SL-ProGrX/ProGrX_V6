VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRH_Cat_Puestos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Catálogo de Puestos"
   ClientHeight    =   6390
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   8160
      TabIndex        =   9
      Top             =   600
      Width           =   1092
      _Version        =   1310722
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5052
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   9612
      _Version        =   1310722
      _ExtentX        =   16954
      _ExtentY        =   8911
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
      Item(0).Caption =   "Puesto"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "Label1(1)"
      Item(0).Control(1)=   "txtDescripcion"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "chkContralaSalario"
      Item(0).Control(4)=   "gbSalario"
      Item(0).Control(5)=   "gbReporta"
      Item(0).Control(6)=   "txtMercadoId"
      Item(0).Control(7)=   "chkControlaMarcas"
      Item(1).Caption =   "Responsabilidades"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "Label1(3)"
      Item(1).Control(1)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4452
         Left            =   -68200
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1310722
         _ExtentX        =   13779
         _ExtentY        =   7853
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox gbReporta 
         Height          =   1092
         Left            =   1800
         TabIndex        =   20
         Top             =   4080
         Width           =   7572
         _Version        =   1310722
         _ExtentX        =   13356
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Reporta a:"
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
         Begin XtremeSuiteControls.FlatEdit txtReportaId 
            Height          =   312
            Left            =   720
            TabIndex        =   21
            Top             =   480
            Width           =   1452
            _Version        =   1310722
            _ExtentX        =   2561
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
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtReportaDesc 
            Height          =   312
            Left            =   2160
            TabIndex        =   22
            Top             =   480
            Width           =   5412
            _Version        =   1310722
            _ExtentX        =   9546
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
      End
      Begin XtremeSuiteControls.GroupBox gbSalario 
         Height          =   1812
         Left            =   1800
         TabIndex        =   13
         Top             =   2280
         Width           =   7452
         _Version        =   1310722
         _ExtentX        =   13144
         _ExtentY        =   3196
         _StockProps     =   79
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtSalarioMaximo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   3480
            TabIndex        =   19
            Top             =   1320
            Width           =   1932
            _Version        =   1310722
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalarioMinimo 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   3480
            TabIndex        =   17
            Top             =   840
            Width           =   1932
            _Version        =   1310722
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalarioActual 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            Height          =   312
            Left            =   3480
            TabIndex        =   15
            Top             =   360
            Width           =   1932
            _Version        =   1310722
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   6
            Left            =   1800
            TabIndex        =   18
            Top             =   1320
            Width           =   1692
            _Version        =   1310722
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Salario Máximo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   5
            Left            =   1800
            TabIndex        =   16
            Top             =   840
            Width           =   1692
            _Version        =   1310722
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Salario Mínimo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   4
            Left            =   1800
            TabIndex        =   14
            Top             =   360
            Width           =   1692
            _Version        =   1310722
            _ExtentX        =   2984
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Salario Actual"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   7452
         _Version        =   1310722
         _ExtentX        =   13144
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
      Begin XtremeSuiteControls.FlatEdit txtMercadoId 
         Height          =   312
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   1932
         _Version        =   1310722
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkContralaSalario 
         Height          =   252
         Left            =   1800
         TabIndex        =   12
         Top             =   2040
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Controla Salario?"
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
      Begin XtremeSuiteControls.CheckBox chkControlaMarcas 
         Height          =   252
         Left            =   1800
         TabIndex        =   23
         Top             =   1680
         Width           =   3012
         _Version        =   1310722
         _ExtentX        =   5313
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Controla Marcas?"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   972
         Index           =   3
         Left            =   -69880
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Seleccione las Responsabilidades vinculadas con este puesto:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   492
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Código de Mercado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
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
      Left            =   3480
      TabIndex        =   1
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
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1212
      _Version        =   1310722
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Puesto"
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
Attribute VB_Name = "frmRH_Cat_Puestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vEdita  As Boolean
Dim vCodigo As String, vPaso As Boolean


Private Function fxExiste(vCodigo As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from RH_PUESTOS where COD_PUESTO =  '" & vCodigo & "' "
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxExiste = False
Else
  fxExiste = True
End If
rs.Close
End Function



Private Sub chkContralaSalario_Click()

If chkContralaSalario.Value = xtpChecked Then
    gbSalario.Enabled = True
Else
    gbSalario.Enabled = False
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 COD_PUESTO from RH_PUESTOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_PUESTO > '" & txtCodigo.Text & "' order by COD_PUESTO asc"
    Else
       strSQL = strSQL & " where COD_PUESTO < '" & txtCodigo.Text & "' order by COD_PUESTO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Cod_Puesto
      Call sbConsulta(txtCodigo.Text)
      
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
vModulo = 23
End Sub


Private Sub Form_Load()
vModulo = 23
 
vEdita = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 950
    .Add , , "Descripción", 6750
End With

Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Call sbLimpiaDatos

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaDatos()

vCodigo = ""

tcMain.Item(0).Selected = True

txtCodigo.Text = ""
txtDescripcion.Text = ""
txtMercadoId.Text = ""

txtReportaId.Text = ""
txtReportaDesc.Text = ""

chkActivo.Value = xtpChecked

chkControlaMarcas.Value = xtpUnchecked
chkContralaSalario.Value = xtpUnchecked

txtSalarioActual.Text = Format(0, "Standard")
txtSalarioMinimo.Text = Format(0, "Standard")
txtSalarioMaximo.Text = Format(0, "Standard")

Call chkContralaSalario_Click

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert RH_PUESTOS_ROL(COD_RESPONSABILIDAD,COD_PUESTO,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RH_PUESTOS_ROL where COD_RESPONSABILIDAD = '" & Item.Text & "' and COD_PUESTO = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
    Call sbLswLlena(txtCodigo.Text)
End If

End Sub

'Private Sub txtCodCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF4 Then
'    frmCntX_ConsultaCuentas.Show vbModal
'    txtCodCuenta = gCuenta
'    txtDescCuenta = fxgCntCuentaDesc(gCuenta)
'End If
'
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescCuenta.SetFocus
'End Sub
'
'Private Sub txtCodCuenta_LostFocus()
'txtCodCuenta = fxgCntCuentaFormato(False, txtCodCuenta)
'txtDescCuenta = fxgCntCuentaDesc(txtCodCuenta)
'txtCodCuenta = fxgCntCuentaFormato(True, txtCodCuenta)
'End Sub



Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" And vEdita = True Then Call sbConsulta(txtCodigo.Text)
End Sub



Private Sub sbLswLlena(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem

On Error GoTo vError

vPaso = True

lsw.ListItems.Clear

strSQL = "select R.COD_RESPONSABILIDAD AS 'CODIGO', R.DESCRIPCION,ISNULL(A.COD_PUESTO,'') AS 'Idx'" _
       & "  from RH_RESPONSABILIDADES R" _
       & "   LEFT JOIN RH_PUESTOS_ROL A ON R.COD_RESPONSABILIDAD = A.COD_RESPONSABILIDAD" _
       & "   AND A.COD_PUESTO = '" & pCodigo & "'" _
       & " WHERE R.ACTIVO = 1" _
       & " order by ISNULL(A.COD_PUESTO,'') desc, R.COD_RESPONSABILIDAD"
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!IdX <> "" Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

 strSQL = "select *" _
        & " from vRH_Puestos " _
        & " where COD_PUESTO = '" & pCodigo & "'"
 Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  tcMain.Item(0).Selected = True
  vEdita = True

  txtCodigo.Text = rs!Cod_Puesto
  vCodigo = rs!Cod_Puesto
  
  txtDescripcion.Text = rs!Descripcion
  chkActivo.Value = rs!ACTIVO
  
  txtMercadoId.Text = rs!Cod_Mercado
  
  chkControlaMarcas.Value = rs!Control_Marcas
  chkContralaSalario.Value = rs!Salario_Control
  
  txtSalarioActual.Text = Format(rs!Salario_Actual, "Standard")
  txtSalarioMinimo.Text = Format(rs!Salario_Minimo, "Standard")
  txtSalarioMaximo.Text = Format(rs!Salario_Maximo, "Standard")
  
  txtReportaId.Text = rs!Reporta_Id
  txtReportaDesc.Text = rs!Reporta_Desc
  
  Call chkContralaSalario_Click
    
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  txtCodigo.Text = ""
  txtCodigo.SetFocus
  Call sbLimpiaDatos
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpiaDatos
        vEdita = False
        txtCodigo.SetFocus
       Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaDatos
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_PUESTO,descripcion from RH_PUESTOS "
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

End Select


End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pReportaId As String


If Trim(txtReportaId.Text) = "" Then
    pReportaId = "Null"
Else
    pReportaId = "'" & txtReportaId.Text & "'"
End If

If fxExiste(txtCodigo.Text) Then
  strSQL = "update RH_PUESTOS set descripcion = '" & Trim(txtDescripcion.Text) & "', COD_MERCADO = '" & Trim(txtMercadoId.Text) & "'" _
         & ",ACTIVO = " & chkActivo.Value & ",CONTROL_MARCAS = " & chkControlaMarcas.Value _
         & ",SALARIO_CONTROL = " & chkContralaSalario.Value _
         & ",SALARIO_ACTUAL = " & CCur(txtSalarioActual.Text) & ", SALARIO_MINIMO = " & CCur(txtSalarioMinimo.Text) _
         & ",SALARIO_MAXIMO = " & CCur(txtSalarioMaximo.Text) _
         & ",COD_PUESTO_REPORTA = " & pReportaId _
         & " where COD_PUESTO = '" & vCodigo & "' "
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Puesto de Trabajo: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert into RH_PUESTOS(COD_PUESTO,descripcion,ACTIVO,COD_MERCADO, COD_PUESTO_REPORTA" _
          & ", CONTROL_MARCAS,SALARIO_CONTROL,SALARIO_ACTUAL,SALARIO_MINIMO,SALARIO_MAXIMO,REGISTRO_USUARIO,REGISTRO_FECHA)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion.Text) & "', " & chkActivo.Value & ",'" & Trim(txtMercadoId.Text) _
          & "'," & pReportaId & "," & chkControlaMarcas.Value & "," & chkContralaSalario.Value _
          & "," & CCur(txtSalarioActual.Text) & "," & CCur(txtSalarioMinimo.Text) & "," & CCur(txtSalarioMaximo.Text) _
          & ",'" & glogon.Usuario & " ',dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Puesto de Trabajo: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Function fxValida()

fxValida = True

If Trim(txtCodigo) = "" Then fxValida = False
If Trim(txtDescripcion) = "" Then fxValida = False

End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcMain.Item(0).Selected = True
    txtDescripcion.SetFocus
End If


If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Puesto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_PUESTO"
   gBusquedas.Orden = "COD_PUESTO"
   gBusquedas.Consulta = "select COD_PUESTO,descripcion from RH_PUESTOS"
   frmBusquedas.Show vbModal
   txtCodigo.Text = gBusquedas.Resultado
   
   tcMain.Item(0).Selected = True
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMercadoId.SetFocus
End Sub


Private Sub txtReportaId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Puesto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_PUESTO"
   gBusquedas.Orden = "COD_PUESTO"
   gBusquedas.Consulta = "select COD_PUESTO,descripcion from RH_PUESTOS"
   gBusquedas.Filtro = " AND COD_PUESTO <> '" & vCodigo & "'"
   frmBusquedas.Show vbModal
   txtReportaId.Text = gBusquedas.Resultado
   txtReportaDesc.Text = gBusquedas.Resultado2
End If

End Sub

