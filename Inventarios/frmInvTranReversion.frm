VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvTranReversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reversiones de Movimientos de Inventarios"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdReversar 
      Caption         =   "&Reversar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4800
      Width           =   4935
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   3600
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   38194
   End
   Begin VB.TextBox txtNotas 
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
      Height          =   555
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   5160
      Width           =   4935
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   1695
      Left            =   2160
      TabIndex        =   7
      Top             =   2280
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   0
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
      ForeColor       =   &H00800000&
      Height          =   1635
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   600
      Width           =   6375
   End
   Begin VB.TextBox txtCodigo 
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
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Traslado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton opt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   8640
      X2              =   0
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle"
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
      Height          =   315
      Index           =   4
      Left            =   2160
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Causa"
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
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Afectación"
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
      Height          =   315
      Index           =   2
      Left            =   2160
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8520
      X2              =   2160
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reversión"
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
      Height          =   260
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2040
      X2              =   2040
      Y1              =   240
      Y2              =   5760
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Boleta"
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
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reversar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmInvTranReversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vScroll As Boolean, vMascara As String
Dim vTipo As String, vTipoInverso As String

Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String
' Verificar si se encuentra procesada
' Verificar la fecha de Afectacion
' Verificar que Exista

vMensaje = ""

strSQL = "select estado from pv_InvTranSac where tipo = '" & vTipo _
       & "' and Boleta = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  vMensaje = vMensaje & " - No se encontró # Boleta " & vbCrLf
Else
 If rs!Estado <> "P" Then
      vMensaje = vMensaje & " - La Boleta consultada no se encuentra procesada..." & vbCrLf
 End If
End If
rs.Close

If Not fxInvPeriodos(dtpFecha.Value) Then vMensaje = vMensaje & vbCrLf & " - El periodo en el que desea realizar el movimiento se encuentra cerrado ..."

If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function

Private Sub cmdReversar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As String, vFecha As Date, vOrigen As String

If Not fxVerifica Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError


strSQL = "select isnull(max(Boleta),0)+1 as Ultimo from pv_InvTranSac where Tipo = '" & vTipoInverso & "'"
Call OpenRecordSet(rs, strSQL)
  xCodigo = Format(rs!ultimo, vMascara)
rs.Close

Select Case vTipoInverso
    Case "R"
      vOrigen = "Requisicion"
    Case "E"
      vOrigen = "Entrada"
    Case "S"
      vOrigen = "Salida"
    Case "T"
      vOrigen = "Traslado"
    Case Else
      vOrigen = "No Ident."
End Select


vFecha = Format(dtpFecha.Value, "yyyy/mm/dd") & " " & Time

strSQL = "insert pv_InvTranSac(Boleta,Tipo,cod_entsal,genera_fecha,documento,notas" _
       & ",genera_user,estado,plantilla,fecha,fecha_sistema,autoriza_fecha,autoriza_user" _
       & ",procesa_fecha,procesa_user) values('" & xCodigo & "','" & vTipoInverso & "','" & fxCodigoCbo(cbo) _
       & "',dbo.MyGetdate(),'Rev." & vCodigo & "','" & txtNotas & "','" & glogon.Usuario & "','P',0,'" _
       & Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss") & "',dbo.MyGetdate(),dbo.MyGetdate(),'" & glogon.Usuario _
       & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Call ConectionExecute(strSQL)

Select Case vTipoInverso
 Case "E", "S"
   strSQL = "insert into pv_invTraDet(Linea,Boleta,Tipo,Cod_Bodega,cod_Producto,Cod_Bodega_destino,cantidad,Precio,despacho)" _
          & " (select Linea,'" & xCodigo & "','" & vTipoInverso & "',Cod_Bodega,cod_Producto,Cod_Bodega_destino,cantidad,Precio,cantidad as Desp" _
          & " From pv_invTraDet Where Tipo = '" & vTipo & "' And Boleta = '" & vCodigo & "')"
   Call ConectionExecute(strSQL)
 Case "T"
   strSQL = "insert into pv_invTraDet(Linea,Boleta,Tipo,Cod_Bodega,cod_Producto,Cod_Bodega_destino,cantidad,Precio,despacho)" _
          & " (select Linea,'" & xCodigo & "','" & vTipoInverso & "',Cod_Bodega_destino,cod_Producto,Cod_Bodega,cantidad,Precio,cantidad as Desp" _
          & " From pv_invTraDet Where Tipo = '" & vTipo & "' And Boleta = '" & vCodigo & "')"
   Call ConectionExecute(strSQL)
End Select


strSQL = "select * from pv_invTraDet where tipo = '" & vTipoInverso & "' and Boleta = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  If vTipoInverso = "T" Then
    'Salida x Traslado
    Call sbInvInventario(rs!cod_producto, rs!cantidad, rs!cod_bodega, xCodigo, vOrigen, vFecha _
        , rs!Precio, 0, 0, "S")
    'Entrada x Traslado
    Call sbInvInventario(rs!cod_producto, rs!cantidad, rs!cod_Bodega_Destino, xCodigo, vOrigen, vFecha _
        , rs!Precio, 0, 0, "E")
  Else
    'Entrada/Salida
    Call sbInvInventario(rs!cod_producto, rs!cantidad, rs!cod_bodega, xCodigo, vOrigen, vFecha _
        , rs!Precio, 0, 0, vTipoInverso)
  End If
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
MsgBox "Reversion de " & vTipo & " , Boleta :" & vCodigo & " realizada con " & vOrigen & " boleta : " & xCodigo, vbInformation

Call sbLimpiaPantalla

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
      txtCodigo = rs!Boleta
      Call txtCodigo_LostFocus
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

Private Sub sbLimpiaPantalla()
 
vCodigo = ""
txtCodigo = ""
txtDetalle = ""
 
lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
 
Select Case True
  Case opt.Item(0).Value 'Entradas
    vTipo = "E"
    vTipoInverso = "S"
    lsw.ColumnHeaders.Add , , "Linea", 640
    lsw.ColumnHeaders.Add , , "Articulo", 1140
    lsw.ColumnHeaders.Add , , "Descripción", 2640
    lsw.ColumnHeaders.Add , , "Bodega", 940
    lsw.ColumnHeaders.Add , , "Cantidad", 840
    lsw.ColumnHeaders.Add , , "Cost/Ud", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Total", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Bodega Desc", 2940
  Case opt.Item(1).Value 'Salidas
    vTipo = "S"
    vTipoInverso = "E"
    lsw.ColumnHeaders.Add , , "Linea", 640
    lsw.ColumnHeaders.Add , , "Articulo", 1140
    lsw.ColumnHeaders.Add , , "Descripción", 2640
    lsw.ColumnHeaders.Add , , "Bodega", 940
    lsw.ColumnHeaders.Add , , "Cantidad", 840
    lsw.ColumnHeaders.Add , , "Cost/Ud", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Total", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Bodega Desc", 2940
  Case opt.Item(2).Value 'Traslados
    vTipo = "T"
    vTipoInverso = "T"
    lsw.ColumnHeaders.Add , , "Linea", 640
    lsw.ColumnHeaders.Add , , "Articulo", 1140
    lsw.ColumnHeaders.Add , , "Descripción", 2640
    lsw.ColumnHeaders.Add , , "Bod/Origen", 1140
    lsw.ColumnHeaders.Add , , "Bod/Destino", 1140
    lsw.ColumnHeaders.Add , , "Cantidad", 840
    lsw.ColumnHeaders.Add , , "Cost/Ud", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Total", 1240, vbRightJustify
    lsw.ColumnHeaders.Add , , "Bodega Origen", 2940
    lsw.ColumnHeaders.Add , , "Bodega Destino", 2940
End Select

Call sbInvESCombo(vTipoInverso, cbo)
txtNotas = ""
dtpFecha.Value = fxFechaServidor


End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 32
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 
 vMascara = "0000000000"
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub opt_Click(Index As Integer)
 Call sbLimpiaPantalla
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError


vCodigo = Format(txtCodigo, vMascara)
lsw.ListItems.Clear

strSQL = "select X.*,(rtrim(C.cod_entsal) + ' - ' + C.descripcion) as Causa" _
       & " from PV_INVTRANSAC X inner join pv_entrada_salida C on X.cod_entsal = C.cod_entsal" _
       & " where X.boleta = '" & vCodigo & "' and X.tipo = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  
  vCodigo = rs!Boleta
  txtCodigo = rs!Boleta
    
  txtDetalle = "Documento ......: " & rs!Documento & vbCrLf
  txtDetalle = txtDetalle & "Causa ..........: " & rs!Causa & vbCrLf
  txtDetalle = txtDetalle & "Fecha ..........: " & Format(rs!fecha, "yyyy/mm/dd hh:mm:ss") & vbCrLf
  txtDetalle = txtDetalle & "Notas ..........: " & rs!notas & vbCrLf
    
  dtpFecha.Value = rs!fecha
  txtNotas = ""
    
  Select Case rs!Estado
    Case "S" 'Solicitada
        txtDetalle = txtDetalle & "Estado .........: Solicitado " & vbCrLf
    Case "A" 'Autorizada
        txtDetalle = txtDetalle & "Estado .........: Autorizada" & vbCrLf
    Case "R" 'Rechazada
        txtDetalle = txtDetalle & "Estado .........: Rechazada" & vbCrLf
    Case "P" 'Procesada
        txtDetalle = txtDetalle & "Estado .........: Procesada" & vbCrLf
  End Select
  
  txtDetalle = txtDetalle & "Generado Por ...: " & rs!genera_user & vbCrLf
  txtDetalle = txtDetalle & "Generado Fecha .: " & rs!genera_fecha & vbCrLf
 
  txtDetalle = txtDetalle & "Autorizado Por .: " & rs!Autoriza_user & vbCrLf
  txtDetalle = txtDetalle & "Autorizado Fecha: " & rs!Autoriza_Fecha & vbCrLf
 
  txtDetalle = txtDetalle & "Procesado Por ..: " & rs!Procesa_user & vbCrLf
  txtDetalle = txtDetalle & "Procesado Fecha : " & rs!Procesa_Fecha & vbCrLf
 
    
  strSQL = "select D.Linea,D.cod_producto,P.descripcion,D.cantidad,B.cod_bodega,B.descripcion as Bodega,D.precio,(D.cantidad * D.precio) as Total" _
         & ",isnull(D.despacho,0) as Despacho,D.cod_bodega_destino,X.descripcion as BodegaD" _
         & " from PV_INVTRADET D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " inner join PV_Bodegas B on D.cod_bodega = B.cod_bodega" _
         & " left join PV_Bodegas X on D.cod_bodega_destino = X.cod_Bodega" _
         & " where D.boleta = '" & rs!Boleta & "' and D.tipo = '" & rs!Tipo & "'" _
         & " order by D.Linea"
   rs.Close
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
         
    Set itmX = lsw.ListItems.Add(, , rs!Linea)
        itmX.SubItems(1) = rs!cod_producto
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = rs!cod_bodega
        
    Select Case vTipo
      Case "E", "S"
        itmX.SubItems(4) = rs!cantidad
        itmX.SubItems(5) = Format(rs!Precio, "Standard")
        itmX.SubItems(6) = Format(rs!Precio * rs!cantidad, "Standard")
        itmX.SubItems(7) = rs!Bodega
        
      Case "T"
        itmX.SubItems(4) = rs!cod_Bodega_Destino
        itmX.SubItems(5) = rs!cantidad
        itmX.SubItems(6) = Format(rs!Precio, "Standard")
        itmX.SubItems(7) = Format(rs!Precio * rs!cantidad, "Standard")
        itmX.SubItems(8) = rs!Bodega
        itmX.SubItems(9) = rs!BodegaD
        
    End Select
    rs.MoveNext
   Loop

Else
   MsgBox "No se encontró la Boleta, verifique...", vbExclamation
End If
rs.Close
   
Exit Sub
vError:
 vCodigo = ""
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
