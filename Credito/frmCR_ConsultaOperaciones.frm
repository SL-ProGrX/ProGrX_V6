VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_ConsultaOperaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Operaciones en Tramite"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   HelpContextID   =   3011
   Icon            =   "frmCR_ConsultaOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   9975
   Begin MSComctlLib.ListView lswBusca 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   3413
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "operacion"
         Text            =   "#Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "codigo"
         Text            =   "Código"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "cedula"
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "fecha"
         Text            =   "Fecha Solicitud"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "monto"
         Text            =   "Monto Sol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Estado Sol"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Estado EC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Proceso"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   7
         Top             =   160
         Width           =   2055
      End
      Begin VB.TextBox txtDescripcion 
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
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   200
         Width           =   2055
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "por Cédula"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   160
         Width           =   1575
      End
      Begin VB.OptionButton optTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "por # Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   160
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# Operación"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   200
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDetalle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   720
      Width           =   9915
   End
End
Attribute VB_Name = "frmCR_ConsultaOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UltimaCedula As String

Private Sub cmdReporte_Click()
With Printer
 Printer.Print txtDetalle
 .NewPage
 .EndDoc
End With
End Sub


Private Sub Form_Load()
 UltimaCedula = ""
End Sub

Private Sub lswBusca_Click()
On Error Resume Next
  Operacion.OperacionConsulta = lswBusca.SelectedItem.Text
  lswBusca.Visible = False
  Call LLenaLista(lswBusca.SelectedItem.Text)
End Sub


Private Sub optTipo_Click(Index As Integer)
Select Case Index
 Case 0 '#operacion
  lblDescripcion.Caption = "# Operación"
  lswBusca.Visible = False
 Case 1 'cedula
  lblDescripcion.Caption = "Cédula "
End Select

txtDescripcion = ""
txtDescripcion.SetFocus


End Sub

Private Sub sbBusqueda()

Select Case True
  Case optTipo(0).Value
    gBusquedas.Consulta = "select id_solicitud,codigo,cedula FROM REG_CREDITOS"
    gBusquedas.Columna = "id_solicitud"
    gBusquedas.Orden = "ID_solicitud"
  Case optTipo(1).Value
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "select cedula,nombre from socios"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "cedula"
End Select

frmBusquedas.Show vbModal
 
txtDescripcion = gBusquedas.Resultado
txtDescripcion.SetFocus

End Sub


Sub sbConsulta()
Dim rsbusca As New ADODB.Recordset, itmX As ListItem
Dim strSQL As String

Select Case True
 Case optTipo(0).Value '#operacion
    
    If IsNumeric(txtDescripcion) Then
        Operacion.OperacionConsulta = txtDescripcion
        Call LLenaLista(txtDescripcion.Text)
    End If
    
 Case optTipo(1).Value '#cedula
  Me.MousePointer = vbHourglass
  lswBusca.Visible = True
  
  If UltimaCedula <> txtDescripcion Then
    UltimaCedula = txtDescripcion
    lswBusca.ListItems.Clear
    
    With rsbusca
     strSQL = "Select id_solicitud,codigo,cedula,fechasol,montosol,estadosol,estado,proceso from " _
            & "REG_CREDITOS where cedula = '" & Trim(txtDescripcion) & "' order by id_solicitud desc" 'and estado is null"
     
     Call OpenRecordSet(rsbusca, strSQL, 0)
    Do While Not .EOF
      Set itmX = lswBusca.ListItems.Add(, , CStr(!Id_Solicitud))
       itmX.SubItems(1) = !Codigo
       itmX.SubItems(2) = !Cedula
       itmX.SubItems(3) = Format(!FechaSol, "yyyy/mm/dd")
       itmX.SubItems(4) = Format(!montosol, "###,###,###,##0.00")
       Select Case !estadosol
        Case "R"
         itmX.SubItems(5) = "Recibida"
        Case "P"
         itmX.SubItems(5) = "Pendiente"
        Case "A"
         itmX.SubItems(5) = "Aprobada"
        Case "D"
         itmX.SubItems(5) = "Denegada"
        Case "F"
         itmX.SubItems(5) = "Formalizada"
        Case "N"
         itmX.SubItems(5) = "Anulada"
       End Select
       
      Select Case !Estado
        Case "A"
         itmX.SubItems(6) = "Activa"
        Case "C"
         itmX.SubItems(6) = "Cancelada"
        Case Else
         itmX.SubItems(6) = "En Tramite"
      End Select
      
      Select Case !Proceso
        Case "J"
         itmX.SubItems(7) = "Cobro Jud"
        Case "N"
         itmX.SubItems(7) = "Normal"
        Case "T"
         itmX.SubItems(7) = "Traspaso"
        Case Else
         itmX.SubItems(7) = "------"
      End Select
      
      
      .MoveNext
     Loop
      .Close
    End With
  End If 'ultima cedula
  Me.MousePointer = vbDefault
End Select
End Sub


Sub LLenaLista(strOperacion As String)
Dim rs As New ADODB.Recordset, itmX As ListItem
Dim strEstado As String, strSQL As String, rs2 As New ADODB.Recordset
Dim strTipoDoc As String, strNumDoc As String, lngId As Long
Dim i As Integer, vSegmento As String


On Error Resume Next

Me.MousePointer = vbHourglass
 
  strSQL = "Select R.*,S.nombre,C.Descripcion as DescCod from reg_creditos R inner join Socios S" _
         & " On R.cedula = S.cedula Inner join Catalogo C on R.codigo = C.codigo" _
         & " Where ID_SOLICITUD = " & strOperacion
  Call OpenRecordSet(rs, strSQL)
  
  
  txtDetalle = "" & vbCrLf & vbCrLf
  
  If rs.EOF And rs.BOF Then
    Me.MousePointer = vbDefault
    MsgBox "NO SE ENCONTRO OPERACION", vbInformation
    rs.Close
    Exit Sub
  End If
  
  vSegmento = "FRM-" & Format(rs!fechaforp, "yyyymmdd")
  
  txtDetalle = txtDetalle & "OPERACION : " & Trim(rs!Id_Solicitud) & vbCrLf
  txtDetalle = txtDetalle & "CEDULA    : " & rs!Cedula & vbTab & rs!Nombre & vbCrLf
  txtDetalle = txtDetalle & "CODIGO    : " & rs!Codigo & vbTab & vbTab & rs!descCod & vbCrLf
  txtDetalle = txtDetalle & "GARANTIA  : " & UCase(fxGarantia(rs!Garantia)) & vbTab _
             & "ESTADO : " & UCase(fxEstadoOperacion(rs!estadosol)) & vbCrLf & vbCrLf

'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  txtDetalle = txtDetalle & "*****************  RECEPCION *****************" & vbCrLf & vbCrLf
  txtDetalle = txtDetalle & "USUARIO : " & rs!userRec & vbTab & "FECHA : " _
             & Format(rs!FechaSol, "dd/mm/yyyy") & vbCrLf
  txtDetalle = txtDetalle & "MONTO   : " & Format(rs!montosol, "###,###,###,##0.00") & vbTab _
             & "PLAZO   : " & rs!Plazo & vbTab & "INT : " & rs!Int & vbCrLf _
             & "CUOTA   : " & rs!Cuota & vbCrLf & vbCrLf
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  txtDetalle = txtDetalle & "***************** RESOLUCION *****************" & vbCrLf & vbCrLf
  txtDetalle = txtDetalle & "USUARIO : " & rs!userres & vbTab & "FECHA : " _
             & Format(rs!fechares, "dd/mm/yyyy") & vbCrLf
  txtDetalle = txtDetalle & "MONTO   : " & Format(rs!montoapr, "###,###,###,##0.00") & vbTab _
             & "PLAZO   : " & rs!Plazo & vbTab & "INT : " & rs!Int & vbCrLf _
             & "CUOTA   : " & rs!Cuota & vbCrLf & vbCrLf
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  txtDetalle = txtDetalle & "***************** FORMALIZACION *****************" & vbCrLf & vbCrLf
  txtDetalle = txtDetalle & "USUARIO : " & rs!Userfor & vbTab & "FECHA : " _
             & Format(rs!fechaforp, "dd/mm/yyyy") & vbCrLf
  txtDetalle = txtDetalle & "FEC.CALCULO : " & Format(rs!fecha_calculo_int, "dd/mm/yyyy") _
             & vbTab & "MONTO GIRADO : " & Format(rs!monto_girado, "###,###,###,##0.00") & vbCrLf
  txtDetalle = txtDetalle & "DOCUMENTO   : " & rs!TDOCUMENTO & "-" & rs!nDocumento & vbCrLf
  txtDetalle = txtDetalle & "DESEMBOLSOS : " & IIf(IsNull(rs!documento_referido), "N/A", rs!documento_referido) & vbCrLf & vbCrLf
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
    
  strSQL = "select F.*,S.nombre as Nomb " _
         & "from fiadores F inner join Socios S on F.cedulaf = S.cedula " _
         & "where F.id_solicitud = " & rs!Id_Solicitud
  rs2.CursorLocation = adUseServer
  rs2.Open strSQL, glogon.Conection, adOpenStatic
  If Not rs2.EOF And Not rs2.BOF Then
      txtDetalle = txtDetalle & "***************** REGISTRO DE FIADORES ***************** " & vbCrLf & vbCrLf
      Do While Not rs2.EOF
        txtDetalle = txtDetalle & "CEDULA : " & rs2!cedulaf & vbTab & "NOMBRE : " & rs2!Nombre & vbCrLf
        rs2.MoveNext
      Loop
  Else
      txtDetalle = txtDetalle & " ** NO EXISTEN FIADORES REGISTRADOS PARA ESTA SOLICITUD ** " & vbCrLf
  End If
  rs2.Close
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  strSQL = "select * from refundiciones " _
         & "where id_solicitudr = " & rs!Id_Solicitud
  rs2.CursorLocation = adUseServer
  rs2.Open strSQL, glogon.Conection, adOpenStatic
  If Not rs2.EOF And Not rs2.BOF Then
      txtDetalle = txtDetalle & "***************** REFUNDICIONES A CARTERA *****************" & vbCrLf & vbCrLf
      Do While Not rs2.EOF
        txtDetalle = txtDetalle & "OPERACION : " & rs2!Id_Solicitud & vbTab & "CODIGO : " _
                   & rs2!Codigo & vbTab & "MONTO : " & Format(rs2!Monto, "###,###,###,##0.00") _
                   & " INTC.Atr." & Format(rs2!IntCor, "###,###,###,##0.00") & vbTab & " INT.MORO : " _
                   & Format(rs2!IntMor, "###,###,###,##0.00") & vbCrLf
        rs2.MoveNext
      Loop
  Else
      txtDetalle = txtDetalle & " ** NO EXISTEN REFUNDICIONES A CARTERA REGISTRADAS PARA ESTA SOLICITUD ** " & vbCrLf
  End If
  rs2.Close
  
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  strSQL = "select * from refunde_retencion " _
         & "where id_solicitudr = " & rs!Id_Solicitud
  rs2.CursorLocation = adUseServer
  rs2.Open strSQL, glogon.Conection, adOpenStatic
  If Not rs2.EOF And Not rs2.BOF Then
      txtDetalle = txtDetalle & "***************** REFUNDICIONES DE RETENCIONES *****************" & vbCrLf & vbCrLf
      Do While Not rs2.EOF
        txtDetalle = txtDetalle & "OPERACION : " & rs2!Id_Solicitud & vbTab & "CODIGO : " _
                   & rs2!Codigo & vbTab & "MONTO : " & Format(rs2!Monto, "###,###,###,##0.00") _
                   & " MOROSIDAD : " & Format(rs2!Mora, "###,###,###,##0.00") & vbCrLf

        rs2.MoveNext
      Loop
  Else
      txtDetalle = txtDetalle & " ** NO EXISTEN REFUNDICIONES DE RETENCIONES REGISTRADAS PARA ESTA SOLICITUD ** " & vbCrLf
  End If
  rs2.Close
  
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  strSQL = "select * from desembolsos " _
         & "where id_solicitud = " & rs!Id_Solicitud
  rs2.CursorLocation = adUseServer
  rs2.Open strSQL, glogon.Conection, adOpenStatic
  If Not rs2.EOF And Not rs2.BOF Then
      txtDetalle = txtDetalle & "***************** DESEMBOLSOS *****************" & vbCrLf & vbCrLf
      Do While Not rs2.EOF
        txtDetalle = txtDetalle & "MONTO : " & Format(rs2!Monto, "###,###,###,##0.00") & vbTab _
                   & "CUENTA : " & Format(rs2!Cuenta_Conta, GLOBALES.gstrMascara) & vbTab _
                   & "BENEFICIARIO : " & rs2!concepto & vbCrLf
        rs2.MoveNext
      Loop
  Else
      txtDetalle = txtDetalle & " ** NO EXISTEN DESEMBOLSOS REGISTRADOS PARA ESTA SOLICITUD ** " & vbCrLf
  End If
  rs2.Close
  
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  txtDetalle = txtDetalle & "***************** DATOS ADICIONALES *****************" & vbCrLf & vbCrLf
  txtDetalle = txtDetalle & "COMITE : " & UCase(fxDescribe(CStr(rs!ID_COMITE), "COMITES")) _
             & vbTab & "ACTA : " & rs!acta & vbCrLf
  txtDetalle = txtDetalle & "OBSERVACIONES : " & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 1, 70))) & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 71, 70))) & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 142, 70))) & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 213, 70))) & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 284, 70))) & vbCrLf
  txtDetalle = txtDetalle & UCase(IIf(IsNull(rs!observacion), "", Mid(rs!observacion, 355, 70))) & vbCrLf
    
  If Not IsNull(rs!Estado) Then
   Select Case rs!Estado
    Case "A"
       txtDetalle = txtDetalle & "ESTE CREDITO SE ENCUENTRA ACTIVO" & vbCrLf
    Case "C"
       txtDetalle = txtDetalle & "ESTE CREDITO SE ENCUENTRA CANCELADO" & vbCrLf
   End Select
  
  End If
  
  strTipoDoc = rs!TDOCUMENTO & ""
  strNumDoc = rs!nDocumento & ""
 
  rs.Close
  
'  txtDetalle = txtDetalle & "----------------------------------------------------------------------------------------" & vbCrLf
  txtDetalle = txtDetalle & vbCrLf & vbCrLf & vbCrLf
  
  txtDetalle = txtDetalle & "***************** DOCUMENTO EMITIDO *****************" & vbCrLf & vbCrLf
    
  If UCase(strTipoDoc) = "TE" Or UCase(strTipoDoc) = "CK" Then
     'Buscar En Tesoreria
     lngId = 0
     strSQL = "select * from Tes_Transacciones where tipo = '" & strTipoDoc & "' AND Ndocumento = '" _
            & strNumDoc & "'"
     rs.CursorLocation = adUseServer
     Call OpenRecordSet(rs, strSQL)
     
     If rs.EOF And rs.BOF Then
       rs.Close
        strSQL = "select * from Tes_Transacciones where tipo = '" & strTipoDoc & "' AND Ndocumento = '" _
               & Mid(strNumDoc, 2, 11) & Right(strNumDoc, 2) & "'"
        rs.CursorLocation = adUseServer
        Call OpenRecordSet(rs, strSQL)
     End If
     
     
     If rs.EOF And rs.BOF Then
        txtDetalle = txtDetalle & " **** NO SE ENCONTRO DOCUMENTO ***"
     Else
        txtDetalle = txtDetalle & " DOCUMENTO : " & strTipoDoc & "-" & strNumDoc & vbCrLf & vbCrLf
        txtDetalle = txtDetalle & "CUENTA" & vbTab & vbTab & vbTab & vbTab & "    DEBITOS" & vbTab & "   CREDITOS" & vbTab & "DESCRIPCION" & vbCrLf & vbCrLf
        strSQL = "select * from Tes_Trans_Asiento where nsolicitud = " & rs!NSolicitud
        rs.Close
        rs.CursorLocation = adUseServer
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
            txtDetalle = txtDetalle & Format(rs!CUENTA_CONTABLE, GLOBALES.gstrMascara) & vbTab
            strSQL = Format(rs!Monto, "###,###,###,##0.00")
            For i = Len(strSQL) To 18
               strSQL = " " & strSQL
            Next i
            If rs!debehaber = "D" Then
               txtDetalle = txtDetalle & vbTab & strSQL & vbTab & vbTab & vbTab
            Else
               txtDetalle = txtDetalle & vbTab & vbTab & vbTab & strSQL & vbTab
            End If
            txtDetalle = txtDetalle & fxgCntCuentaDesc(rs!CUENTA_CONTABLE) & vbCrLf
          rs.MoveNext
        Loop
        rs.Close
     End If
     rs.Close
  End If

  If UCase(strTipoDoc) = "ND" And Len(Trim(strNumDoc)) > 0 Then
    'Formalizaciones en ASE Documentos
    txtDetalle = txtDetalle & " DOCUMENTO : " & strTipoDoc & "-" & strNumDoc & vbCrLf & vbCrLf
    txtDetalle = txtDetalle & "CUENTA" & vbTab & vbTab & vbTab & vbTab & "    DEBITOS" & vbTab & "   CREDITOS" & vbTab & "DESCRIPCION" & vbCrLf & vbCrLf
    strSQL = "select A.*,B.Descripcion " _
           & " from ase_asientos A inner join cuentas B on A.recas_cuenta = B.cod_cuenta" _
           & " where tipo = '" & strTipoDoc & "' AND id_documento = " & strNumDoc
    rs.CursorLocation = adUseServer
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      txtDetalle = txtDetalle & Format(rs!recas_cuenta, GLOBALES.gstrMascara) & vbTab
                 
      strSQL = Format(rs!recas_monto, "###,###,###,##0.00")
      For i = Len(strSQL) To 18
         strSQL = " " & strSQL
      Next i
      If rs!recas_debehaber = "D" Then
         txtDetalle = txtDetalle & vbTab & strSQL & vbTab & vbTab & vbTab
      Else
         txtDetalle = txtDetalle & vbTab & vbTab & vbTab & strSQL & vbTab
      End If
      txtDetalle = txtDetalle & rs!Descripcion & vbCrLf
      rs.MoveNext
    Loop
    rs.Close
        
    
  End If


  If UCase(strTipoDoc) = "FR" And Len(Trim(strNumDoc)) > 0 Then
    'Formalizaciones en ASE Documentos
    txtDetalle = txtDetalle & " DOCUMENTO : " & strTipoDoc & "-" & strNumDoc & vbCrLf
    txtDetalle = txtDetalle & " SEGMENTO  : " & vSegmento & vbCrLf & vbCrLf
    txtDetalle = txtDetalle & "CUENTA" & vbTab & vbTab & vbTab & vbTab & "    DEBITOS" & vbTab & "   CREDITOS" & vbTab & "DESCRIPCION" & vbCrLf & vbCrLf
    strSQL = "select A.*,B.Descripcion " _
           & " from asientos_TMP A inner join cuentas B on A.TMP_cuenta = B.cod_cuenta" _
           & " where TMP_tipo = 'FRM' AND TMP_OPERACION = " & strNumDoc
    rs.CursorLocation = adUseServer
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      txtDetalle = txtDetalle & Format(rs!tmp_cuenta, GLOBALES.gstrMascara) & vbTab
                 
      strSQL = Format(rs!tmp_monto, "###,###,###,##0.00")
      For i = Len(strSQL) To 18
         strSQL = " " & strSQL
      Next i
      If rs!tmp_debehaber = "D" Then
         txtDetalle = txtDetalle & vbTab & strSQL & vbTab & vbTab & vbTab
      Else
         txtDetalle = txtDetalle & vbTab & vbTab & vbTab & strSQL & vbTab
      End If
      txtDetalle = txtDetalle & rs!Descripcion & vbCrLf
      rs.MoveNext
    Loop
    rs.Close
        
    
  End If


Me.MousePointer = vbDefault

End Sub


Function fxDescribe(strCodigo As String, Tabla As String) As String
Dim rsDes As New ADODB.Recordset, Resultado As String
On Error Resume Next

Select Case Tabla
 Case "REG_CREDITOS"
   rsDes.Source = "Select descripcion as describe from catalogo where codigo = '" & Trim(strCodigo) & "'"
 Case "SOCIOS"
   rsDes.Source = "Select nombre as describe from socios where cedula = '" & Trim(strCodigo) & "'"
 Case "COMITES"
   rsDes.Source = "Select descripcion as describe from comites where id_comite = " & Trim(strCodigo)
End Select

rsDes.ActiveConnection = glogon.Conection
rsDes.Open

Resultado = IIf(IsNull(rsDes!describe), "", rsDes!describe)
rsDes.Close

fxDescribe = Resultado

End Function

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtDescripcion = xOpTemp
 Call txtDescripcion_KeyDown(vbKeyReturn, 0)
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then sbConsulta
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub
