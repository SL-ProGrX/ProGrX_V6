VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCntX_ERConfiguracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configura Estado de Resultados / Excedentes / Utilidades"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   HelpContextID   =   17
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsw 
      Height          =   2412
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   492
      Left            =   6720
      TabIndex        =   4
      Top             =   3720
      Width           =   1572
      _Version        =   1310723
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_ERConfiguracion.frx":0000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleccione los tipos de cuentas que pertenecen a esta sección"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   120
      TabIndex        =   3
      Top             =   696
      Width           =   8292
   End
   Begin VB.Label Label1 
      Caption         =   "Secciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmCntX_ERConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub sbCargaLsw()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListItem

strSQL = "select * from CntX_Tipos_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta

Select Case cbo.Text
  Case "OTROS INGRESOS"
    strSQL = strSQL & " and clasificacion = 'I' and (er is null or er = 'OI')"
  Case "OTROS GASTOS"
    strSQL = strSQL & " and clasificacion = 'G' and (er is null or er = 'OG')"
  Case "VENTAS"
    strSQL = strSQL & " and clasificacion = 'I' and (er is null or er = 'VE')"
  Case "COSTO VENTAS"
    strSQL = strSQL & " and clasificacion = 'G' and (er is null or er = 'CV')"
End Select

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!tipo_cuenta)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!eR) Then itmX.Checked = True
  rs.MoveNext
Loop
rs.Close

End Sub


Private Sub cbo_Click()
 Call sbCargaLsw
End Sub

Private Sub cmdAplicar_Click()
Dim strSQL As String, lng As Long, vTipo As String


Select Case cbo.Text
  Case "OTROS INGRESOS"
     vTipo = "OI"
  Case "OTROS GASTOS"
     vTipo = "OG"
  Case "VENTAS"
     vTipo = "VE"
  Case "COSTO VENTAS"
     vTipo = "CV"
End Select

For lng = 1 To lsw.ListItems.Count
 lsw.SelectedItem = lsw.ListItems(lng)
 If lsw.SelectedItem.Checked Then
   strSQL = "update CntX_Tipos_Cuentas set ER = '" & vTipo & "' where cod_contabilidad = " _
          & gCntX_Parametros.CodigoConta & " and tipo_cuenta = '" _
          & lsw.SelectedItem & "'"
 Else
   strSQL = "update CntX_Tipos_Cuentas set ER = Null where cod_contabilidad = " _
          & gCntX_Parametros.CodigoConta & " and tipo_cuenta = '" _
          & lsw.SelectedItem & "'"
 End If
 Call ConectionExecute(strSQL, 0)
Next lng

MsgBox "Cambios Aplicados ...", vbInformation


End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset

Set Me.Icon = frmContenedor.Icon


Call Formularios(Me)
Call RefrescaTags(Me)

rs.Open "select razonsocial from CntX_Contabilidades where cod_contabilidad = " _
        & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

cbo.AddItem "OTROS INGRESOS"
cbo.AddItem "OTROS GASTOS"

'If rs!razonsocial = "C" Then
'    cbo.AddItem "VENTAS"
'    cbo.AddItem "COSTO VENTAS"
'End If

rs.Close

cbo.Text = "OTROS INGRESOS"

Call sbCargaLsw
End Sub


