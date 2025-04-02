VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCxC_FacturaEstados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descuento de Facturas: Estado de la Factura"
   ClientHeight    =   6408
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   10452
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6408
   ScaleWidth      =   10452
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10092
      _Version        =   524288
      _ExtentX        =   17801
      _ExtentY        =   8700
      _StockProps     =   64
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
      MaxCols         =   485
      ScrollBars      =   2
      SpreadDesigner  =   "frmCxC_FacturaEstados.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Factura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1880
      TabIndex        =   1
      Top             =   360
      Width           =   6855
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCxC_FacturaEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
Dim strSQL As String

Set Me.Icon = frmContenedor.Icon

vModulo = 31

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select Factura_Estado,descripcion,Proceso,Accion,activo from CXC_FACTURAS_ESTADOS" _
       & " order by Factura_Estado"
Call sbCargaGrid(vGrid, 5, strSQL)

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

fxGuardar = 0

vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from CXC_FACTURAS_ESTADOS" _
       & " where Factura_Estado = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into CXC_FACTURAS_ESTADOS(Factura_Estado,descripcion,Proceso,Accion,activo,registro_fecha, registro_usuario)" _
         & " values('" & vGrid.Text
  vGrid.col = 2
  strSQL = strSQL & "','" & vGrid.Text
  vGrid.col = 3
  strSQL = strSQL & "','" & vGrid.Text
  vGrid.col = 4
  strSQL = strSQL & "','" & vGrid.Text
  vGrid.col = 5
  strSQL = strSQL & "'," & vGrid.Value & ", dbo.mygetdate(), '" & glogon.Usuario & "' )"
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  
  Call Bitacora("Registra", "Estado de Factura CxC: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.col = 2
    strSQL = "update CXC_FACTURAS_ESTADOS set descripcion = '" & vGrid.Text & "', Proceso = '"
    vGrid.col = 3
    strSQL = strSQL & vGrid.Text & "', Accion = '"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "', Activo = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where Factura_Estado = '" & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
    vGrid.col = 1
   
    Call Bitacora("Modifica", "Estado de Factura CxC: " & vGrid.Text)
 
End If

rs.Close
fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long


On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        
        strSQL = "delete CXC_FACTURAS_ESTADOS where Factura_Estado = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Estado de Factura CxC: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub
