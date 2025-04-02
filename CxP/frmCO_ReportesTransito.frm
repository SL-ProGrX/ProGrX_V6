VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCO_ReportesTransito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Morosidad en Real/Transito"
   ClientHeight    =   5460
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10668
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   10668
   Begin VB.TextBox txtHasta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   22
      Text            =   "80"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtDesde 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   21
      Text            =   "1"
      Top             =   4800
      Width           =   615
   End
   Begin VB.CheckBox chkRepResumen 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   9360
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ComboBox cboOficina 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   4332
   End
   Begin VB.ComboBox cboRecurso 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3960
      Width           =   4332
   End
   Begin VB.CheckBox chkLineas 
      Appearance      =   0  'Flat
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6120
      TabIndex        =   6
      Top             =   3012
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.ComboBox cboCartera 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2280
      Width           =   4332
   End
   Begin VB.ComboBox cboGarantia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   4332
   End
   Begin VB.ComboBox cboDestino 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3600
      Width           =   4332
   End
   Begin VB.TextBox txtReporteX 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6120
      TabIndex        =   2
      ToolTipText     =   "Presione F4 Para Consultar"
      Top             =   3276
      Width           =   975
   End
   Begin VB.ComboBox cboRep 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   4332
   End
   Begin VB.ComboBox cboInstitucion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   276
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2640
      Width           =   4332
   End
   Begin MSComctlLib.ListView lswRepGen 
      Height          =   3732
      Left            =   360
      TabIndex        =   9
      Top             =   1440
      Width           =   4332
      _ExtentX        =   7641
      _ExtentY        =   6583
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reporte"
         Object.Width           =   7128
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   1
      Left            =   4800
      TabIndex        =   26
      Top             =   3240
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Morosidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Index           =   11
      Left            =   2280
      TabIndex        =   25
      Top             =   360
      Width           =   4572
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   0
      Left            =   7440
      TabIndex        =   24
      Top             =   4800
      Width           =   612
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   312
      Index           =   1
      Left            =   6240
      TabIndex        =   23
      Top             =   4800
      Width           =   612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuotas de Atraso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   792
      Index           =   15
      Left            =   6120
      TabIndex        =   20
      Top             =   4440
      Width           =   2652
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   22
      Left            =   4800
      TabIndex        =   19
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Image imgReporteGeneral 
      Height          =   384
      Left            =   10080
      Picture         =   "frmCO_ReportesTransito.frx":0000
      Top             =   4788
      Width           =   384
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Recurso"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   19
      Left            =   4800
      TabIndex        =   18
      Top             =   3960
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cartera"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   18
      Left            =   4800
      TabIndex        =   17
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Garantía"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   17
      Left            =   4800
      TabIndex        =   16
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   16
      Left            =   4800
      TabIndex        =   15
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Label lblXDescribe 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   7080
      TabIndex        =   13
      Top             =   3276
      Width           =   3372
   End
   Begin VB.Label lblRepGen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   4800
      TabIndex        =   11
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCO_ReportesTransito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtReporteX.Enabled = False
 
  strSQL = "select cod_destino + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos"
  Call sbLlenaCbo(cboDestino, strSQL)
  
   strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_grupos"
   Call sbLlenaCbo(cboRecurso, strSQL)
  
Else
  txtReporteX.Enabled = True

  strSQL = "select (R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtReporteX & "'"
  Call sbLlenaCbo(cboDestino, strSQL)

  strSQL = "select (R.cod_grupo) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtReporteX & "'"
  Call sbLlenaCbo(cboRecurso, strSQL)


End If

End Sub


Private Sub txtReporteX_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF4 Then Call sbBusqueda(7)
If KeyCode = vbKeyReturn Then cboDestino.SetFocus
End Sub

Private Sub txtReporteX_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
End If
End Sub



Private Sub txtReporteX_LostFocus()
 If Len(Trim(txtReporteX)) > 0 Then lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
 Call chkLineas_Click

End Sub

Private Sub Form_Load()


Dim strSQL As String, rs As New ADODB.Recordset


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

 
strSQL = "select rtrim(cod_estado) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from afi_estados_persona"
Call sbLlenaCbo(cboRep, strSQL, True, False)
'Item Adicional
cboRep.AddItem "X - Ex.Socios"
 

 
strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
       & " from instituciones"
Call sbLlenaCbo(cboInstitucion, strSQL, True, True)


 
 With lswRepGen.ListItems
   .Clear
   .Add , "GENDET", "Listado General"
   .Add , "GENAGD", "Listado General Agrupado"
   
   .Add , "ESPCON", "Especial Convenios"
   .Add , "MORCAR", "Comparativo - Resumen"
   
   'Garantia
   .Add , "MORGAR", "Mora x Garantía"
   .Add , "MORGAG", "Mora x Garantía Agrupado"
 
   'Antiguedad Saldos
   .Add , "ANTSALDOS", "Antiguedad de Saldos"
   .Add , "ANTSALDOSPA", "Antiguedad de Saldos + (Producto Acumulado)"
   .Add , "ANTPROACUM", "Antiguedad de Producto Acumulado"
 
 
   'Antiguedad Saldos (Efectos Moratorios y Cobrabilidad)
   .Add , "ANTLEGAL", "Antiguedad de Saldos (Legal)"
   .Add , "ANTFINAN", "Antiguedad de Saldos (Financiera)"
 
   'Nuevos
   .Add , "DETPRV", "Listado x Provincia"
   .Add , "DETPVT", "Listado x Provincia Trabajo"
   .Add , "DETUND", "Listado x Unidad"
   .Add , "LSTFRM", "Listado x Formalización"
   .Add , "LSTCMT", "Listado x Comité Resolutor"
 
 
 
 End With
 
 cboDestino.Clear
 cboDestino.AddItem "TODOS"
 cboDestino.Text = "TODOS"
 
 cboRecurso.Clear
 cboRecurso.AddItem "TODOS"
 cboRecurso.Text = "TODOS"
 
 
strSQL = "select rtrim(Garantia) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from crd_garantia_tipos"
Call sbLlenaCbo(cboGarantia, strSQL, True, False)

strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as Itmx" _
       & " from SIF_Oficinas"
Call sbLlenaCbo(cboOficina, strSQL, True, False)

 
 vPaso = True
 
 cboCartera.Clear
 cboCartera.AddItem "(Todas las Carteras)"
 
 strSQL = "select rtrim(cod_clasificacion) + ' - ' + descripcion as ItemX" _
        & " from CBR_CLASIFICACION_CARTERA order by cod_clasificacion"
 Call OpenRecordSet(rs, strSQL)
 Do While Not rs.EOF
  cboCartera.AddItem rs!itemx
  rs.MoveNext
 Loop
 cboCartera.Text = "(Todas las Carteras)"
 rs.Close
 
 vPaso = False

'Tab OP Generadas


 
 
End Sub

Private Sub lswRepGen_ItemClick(ByVal Item As MSComctlLib.ListItem)

lblRepGen.Caption = Item.Text
lblRepGen.Tag = Item.Key

End Sub

Private Sub imgReporteGeneral_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime Reportes Generales de Cobro
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Utiliza variables globales
'-------------------------------------------------------------------------------------------
Dim strSQL As String, vSubTitulo As String
Dim i As Byte

Me.MousePointer = vbHourglass

strSQL = ""

If cboRep.Text <> "TODOS" Then
    Select Case SIFGlobal.fxCodText(cboRep.Text)
     Case "X"  'Todos
       strSQL = "({SOCIOS.ESTADOACTUAL} = 'A' OR {SOCIOS.ESTADOACTUAL} = 'P')"
     Case Else
       strSQL = "{SOCIOS.ESTADOACTUAL} = '" & SIFGlobal.fxCodText(cboRep.Text) & "'"
    End Select
End If

If cboOficina.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_OFICINA_R} = '" & SIFGlobal.fxCodText(cboOficina.Text) & "'"
End If

If cboInstitucion.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If



If cboDestino.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_DESTINO} = '" & SIFGlobal.fxCodText(cboDestino.Text) & "'"
End If

If cboRecurso.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.COD_GRUPO} = '" & SIFGlobal.fxCodText(cboRecurso.Text) & "'"
End If

If cboGarantia.Text <> "TODOS" Then
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.GARANTIA} = '" & SIFGlobal.fxCodText(cboGarantia.Text) & "'"
End If

vSubTitulo = "Cartera: " & cboCartera.Text & " ¦  Estado: " & cboRep.Text _
                 & " ¦ Garantía: " & cboGarantia.Text & " ¦  Destino: " & cboDestino.Text _
                 & " ¦ Recurso: " & cboRecurso.Text & " ¦ Oficina: " & cboOficina.Text _
                 & " ¦ Institución : " & cboInstitucion.Text


If chkLineas.Value = vbChecked Then
  vSubTitulo = vSubTitulo & " / Línea : Todas"
Else
  If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
  strSQL = strSQL & "{REG_CREDITOS.CODIGO} = '" & txtReporteX.Text & "'"
  vSubTitulo = vSubTitulo & " / Línea : " & txtReporteX.Text
End If

If cboCartera.Text <> "(Todas las Carteras)" Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    strSQL = strSQL & "{CBR_CLASIFICACION_DETALLE.COD_CLASIFICACION} = '" & SIFGlobal.fxCodText(cboCartera.Text) & "'"
End If


vSubTitulo = Mid(vSubTitulo, 1, 250)

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"
     
    .Connect = glogon.ConectRPT
     
  Select Case lblRepGen.Tag
   Case "GENDET" 'Reporte General
    
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumen.rpt")
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD'"
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetallado.rpt")
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD'"
        End If
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL

   
   Case "GENAGD"    ' General - Detallado Agrupado"
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD - ESP.AGRUPADO'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenAgr.rpt")
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD - ESP.AGRUPADO'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoAgr.rpt")
        End If
            
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
   
   Case "ESPCON" 'Convenios
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXConvenios.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & " / FILTRO " & Mid(cboRep.Text, 4, 30) & "'"
        .Formulas(3) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta & " And {VISTA_MOROSIDAD.CODIGO}='" & txtReporteX & "'"
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
  
   Case "MORCAR" 'Resumen Comparativo
        
        
        If chkRepResumen.Value = vbChecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ComparativoRsm.rpt")
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Comparativo.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        Select Case Mid(cboRep.Text, 1, 1)
         Case "T" 'Todos
            .StoredProcParam(0) = "."
            .StoredProcParam(1) = "."
            .StoredProcParam(2) = "."
            .StoredProcParam(3) = "."
         Case "P", "A", "X" 'Opex
            .StoredProcParam(0) = "A"
            .StoredProcParam(1) = "A"
            .StoredProcParam(2) = "P"
            .StoredProcParam(3) = "P"
         Case Else
            .StoredProcParam(0) = SIFGlobal.fxCodText(cboRep.Text)
            .StoredProcParam(1) = SIFGlobal.fxCodText(cboRep.Text)
            .StoredProcParam(2) = SIFGlobal.fxCodText(cboRep.Text)
            .StoredProcParam(3) = SIFGlobal.fxCodText(cboRep.Text)
        End Select
         
       strSQL = ""
       If cboGarantia.Text <> "TODOS" Then
          strSQL = "{spCBRComparativo;1.garantia} = '" & SIFGlobal.fxCodText(cboGarantia.Text) & "'"
       End If
       
       If chkLineas.Value = vbUnchecked Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{spCBRComparativo;1.codigo} = '" & txtReporteX.Text & "'"
       End If
    
       If cboCartera.Text <> "(Todas las Carteras)" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{spCBRComparativo;1.cod_clasificacion} = '" & SIFGlobal.fxCodText(cboCartera.Text) & "'"
       End If
     
     
       .SelectionFormula = strSQL
     
  
   Case "MORGAR" 'Reporte Mora x Garantia
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXGarantia.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        If Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
  
    Case "MORGAG"  'Mora x Garantía - Agrupado
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoXGarantiaAgr.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        If Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = "{REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'" & strSQL
  
  
   Case "DETPRV" 'Provincia
        
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x PROVINCIA'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenProv.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x PROVINCIA '"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoProv.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
   

  Case "DETPVT" 'Detalle x Provincia Trabajo
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x PROVINCIA (Laboral)'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenProvTra.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x PROVINCIA (Laboral)'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoProvTra.rpt")
        End If
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
  
    
    Case "DETUND" 'Detalle x Unidad
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x UNIDAD'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoResumenUnidad.rpt")
            Me.MousePointer = vbDefault
            i = MsgBox("Desea Mostrar x Líneas de Crédito", vbYesNo)
            If i = vbYes Then
              .Formulas(5) = "fxResumen=0"
            Else
              .Formulas(5) = "fxResumen=1"
            End If
        
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x UNIDAD'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoDetalladoUnidad.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
    
    Case "LSTFRM" 'Listado x Formalización

            .Formulas(1) = "Titulo='LISTADO DE MORA. ANALISIS DE FORMALIZACION'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoFormalizacion.rpt")

        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
    Case "LSTCMT" 'Listado x Comité Resolutor
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='GENERAL RESUMEN DE MOROSIDAD x COMITE'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoComiteRsm.rpt")
        Else
            .Formulas(1) = "Titulo='GENERAL DETALLADO DE MOROSIDAD x COMITE'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoComiteDet.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .Formulas(4) = "Cuotas='Cuotas Desde:" & txtDesde & " Hasta:" & txtHasta & "'"
        .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA}>=" & txtDesde & " And {VISTA_MOROSIDAD.CUOTA}<=" & txtHasta
        If Len(.SelectionFormula) > 0 And Len(strSQL) > 0 Then strSQL = " AND " & strSQL
        .SelectionFormula = .SelectionFormula & strSQL
    
     
     Case "ANTLEGAL", "ANTSALDOS", "ANTSALDOSPA", "ANTPROACUM" 'Antiguedad Legal
     
        Select Case lblRepGen.Tag
           Case "ANTLEGAL"      'Antiguedad de Saldos + Mora Legal
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA LEGAL] [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadLegalRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA LEGAL] [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadLegalDet.rpt")
                End If
           
           Case "ANTSALDOS"     'Antiguedad de Saldos (Pura)
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosDet.rpt")
                End If
           
           Case "ANTSALDOSPA"   'Antiguedad de Saldos + (Producto Acumulado)
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosProdAcumRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadSaldosProdAcumDet.rpt")
                End If
           
           Case "ANTPROACUM"    'Antiguedad Producto Acumulado
                If chkRepResumen.Value = vbChecked Then
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [RESUMEN]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadProdAcumRsm.rpt")
                Else
                    .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS + PRODUCTO ACUMULADO [DETALLE]'"
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadProdAcumDet.rpt")
                End If
        
        End Select

        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        
        i = MsgBox("Desea Mostrar Operaciones al Día", vbYesNo)
        If i = vbNo Then
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vCBRAntiguedadSaldos.CAl Día} = 0"
        End If
        
        .SelectionFormula = strSQL
     
     
     
     Case "ANTFINAN" 'Antiguedad Financiera
        If chkRepResumen.Value = vbChecked Then
            .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA FINANCIERA] [RESUMEN]'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadFinancieraRsm.rpt")
        Else
            .Formulas(1) = "Titulo='ANTIGUEDAD DE SALDOS [MORA FINANCIERA] [DETALLE]'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_ListadoAntiguedadFinancieraDet.rpt")
        End If
        
        
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(3) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = strSQL
    
'        i = MsgBox("Desea Mostrar Operaciones al Día", vbYesNo)
'        If i = vbNo Then
'           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'           strSQL = strSQL & "{vCBRAntiguedadSaldos.CAl Día} = 0"
'        End If
  
  End Select
     
     

    .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

