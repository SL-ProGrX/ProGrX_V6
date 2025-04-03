VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_BeneficioPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Pagos de Beneficios Asignados"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsw 
      Height          =   5892
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   10452
      _ExtentX        =   18441
      _ExtentY        =   10398
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Beneficio"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Emite"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "# Cuenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Banco"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "IDBanco"
         Object.Width           =   1658
      EndProperty
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "&Todos"
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   372
      Left            =   7680
      TabIndex        =   6
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   5412
   End
   Begin XtremeSuiteControls.PushButton cmdGenerar 
      Height          =   372
      Left            =   9000
      TabIndex        =   5
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Generar"
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
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   150
      Left            =   0
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   10815
      _Version        =   1441793
      _ExtentX        =   19076
      _ExtentY        =   265
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      Caption         =   "Seleccione los Beneficios (Solicitados para Envio a Bancos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10452
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
End
Attribute VB_Name = "frmAF_BeneficioPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Private Sub cbo_Click()
lsw.ListItems.Clear
chkTodos.Value = 0
End Sub

Private Sub chkTodos_Click()

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, vNombre As String
Dim itmX As ListItem, vTipo As String, vCedula As String
Dim A
lsw.ListItems.Clear
chkTodos.Value = 0
'strSQL = "Select * from afi_bene_pago where cod_beneficio = '" & fxCodigoCbo(cbo) & "'" _
        & " and estado = 'S'"
        
        
strSQL = "Select B.*,U.DESCRIPCION from afi_bene_pago B inner join " _
        & " SOCIOS S on  B.CEDULA = S.CEDULA " _
        & " left join UPROGRAMATICA U on S.UP = U.CODIGO" _
        & " where B.cod_beneficio = '" & SIFGlobal.fxCodText(cbo.Text) & "' and B.ESTADO = 'S' " _
        & " order by U.DESCRIPCION,B.CEDULA"
        
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF

    Select Case rs!Tipo
       Case "S"
         vTipo = "Socio"
         vCedula = rs!Cedula
       Case "B"
         vTipo = "Beneficiario"
    End Select
    'carga la lista
    
    Set itmX = lsw.ListItems.Add(, , rs!consec)
        itmX.SubItems(1) = SIFGlobal.fxCodText(cbo.Text)
        itmX.SubItems(2) = vTipo
        
        vNombre = fxNombre(rs!Cedula)
        'en caso de que sea un beneficiario
        If Trim(vNombre) = "" Or UCase(vNombre) = "ASEASECCSS" Then
           vNombre = fxBeneficiario(rs!Cedula, rs!consec)
        End If
        
        itmX.SubItems(3) = Format(rs!MONTO, "Standard")
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = vNombre
        itmX.SubItems(6) = fxTipoDocumento(rs!Tipo_Emision)
        itmX.SubItems(7) = rs!cta_bancaria & ""
        
        itmX.SubItems(8) = fxDescribeBanco(rs!cod_banco)
        itmX.SubItems(9) = rs!cod_banco
    rs.MoveNext
Loop

rs.Close

End Sub

Private Sub cmdGenerar_Click()
Dim y As Integer, vTesoreria As Long, vMonto As Currency, vIDBanco As Long
Dim vNombre As String, vConsec As Integer, vTipo As String
Dim vBeneficio As String, vCedula As String, vEmite As String
Dim vCtaBanco As String, vCta As String, vBanco As String
Dim vDetalle As String, strSQL As String, vDetalle2 As String
Dim vCtaBene As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked Then y = y + 1
Next i

If y > 0 Then

    PrgBar.Max = y
    PrgBar.Value = 1
    
    strSQL = "select cod_cuenta  from afi_beneficios where cod_beneficio = '" & SIFGlobal.fxCodText(cbo.Text) & "' "
    Call OpenRecordSet(rs, strSQL)
    vCtaBene = IIf(Not IsNull(rs!cod_cuenta), rs!cod_cuenta, "0")
    
   For i = 1 To lsw.ListItems.Count
    
    If lsw.ListItems.Item(i).Checked Then
        
        Select Case UCase(lsw.ListItems.Item(i).SubItems(2))
          Case "SOCIO"
            vTipo = "S"
          Case "BENEFICIARIO"
            vTipo = "B"
        End Select
          
          vConsec = lsw.ListItems.Item(i).Text
          vBeneficio = lsw.ListItems.Item(i).SubItems(1)
          vMonto = CCur(lsw.ListItems.Item(i).SubItems(3))
          vCedula = lsw.ListItems.Item(i).SubItems(4)
          vNombre = Trim(lsw.ListItems.Item(i).SubItems(5))
          vEmite = fxTipoDocumento(lsw.ListItems.Item(i).SubItems(6))
          vIDBanco = lsw.ListItems.Item(i).SubItems(9)
          vBanco = lsw.ListItems.Item(i).SubItems(8)
          vCtaBanco = fxgCtaBanco(vIDBanco)
          vCta = lsw.ListItems.Item(i).SubItems(7)
          
          If vCta <> "" Then
             vCta = fxgCntCuentaFormato(False, vCta)
          End If
          
          vDetalle = Mid(cbo.Text, 1, 27)
          vDetalle2 = Mid(cbo.Text, 28, Len(cbo))
          vTesoreria = fxgTesoreriaMaestro(vEmite, vIDBanco, vMonto, vCedula, vNombre, _
                                           0, vDetalle, 0, vDetalle2, vCta, fxFechaServidor)
          'Actualiza el estado en tabla afi_bene_otorga
          strSQL = "Update afi_bene_otorga set estado = 'E',autoriza_user = '" & glogon.Usuario & "'," _
                  & "autoriza_fecha = dbo.MyGetdate() where cedula = '" & vCedula & "'" _
                  & " and cod_beneficio = '" & vBeneficio & "' and consec = '" & vConsec & "'"
                  
          Call ConectionExecute(strSQL)
          
          'Actualiza estado en afi_bene_pago
          strSQL = "Update afi_bene_pago set estado = 'E',tesoreria = " & vTesoreria & "," _
                  & "envio_user = '" & glogon.Usuario & "',envio_fecha = dbo.MyGetdate() where cedula = '" & vCedula & "'" _
                  & " and cod_beneficio = '" & vBeneficio & "' and consec = '" & vConsec & "'"
                  
          Call ConectionExecute(strSQL)
         
          'Detalle de tesoreria
        
          Call sbgTesoreriaDetalle(vTesoreria, vCtaBanco, vMonto, "H", 1)
          Call sbgTesoreriaDetalle(vTesoreria, vCtaBene, vMonto, "D", 2)
          
    End If
    
    If PrgBar.Max < y Then PrgBar.Value = PrgBar.Value + 1
   
   Next i
  
  
End If

PrgBar.Value = 0

Call cmdBuscar_Click

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
    vModulo = 7
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 7

   imgBanner.Picture = frmContenedor.imgBanner_01.Picture

   strSQL = "select rtrim(cod_Beneficio) + ' - ' + descripcion as Beneficio from afi_beneficios" _
              & " where estado = 'A' and cod_beneficio in (select cod_beneficio from afi_bene_pago" _
              & " where Estado = 'S')"
   Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
     cbo.AddItem rs!Beneficio
     rs.MoveNext
  Loop
  If rs.RecordCount > 0 Then
     rs.MoveFirst
     cbo.Text = rs!Beneficio
  End If
  
  rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxBeneficiario(vCedulaB As String, vConsec As Long) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select nombre from beneficiarios where cedula in(select cedula from afi_bene_pago where cod_beneficio = '" & SIFGlobal.fxCodText(cbo.Text) & "' " _
        & " and consec = " & vConsec & " )and cedulabn = '" & vCedulaB & "'"
       Call OpenRecordSet(rs, strSQL)

If rs.EOF And rs.BOF Then
 fxBeneficiario = ""
Else
 fxBeneficiario = IIf(IsNull(rs!Nombre), "", rs!Nombre)
End If
rs.Close
End Function


