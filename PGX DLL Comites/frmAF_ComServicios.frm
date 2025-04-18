VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_ComServicios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Servicios"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleMode       =   0  'User
   ScaleWidth      =   7608.564
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      _Version        =   524288
      _ExtentX        =   11880
      _ExtentY        =   9128
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_ComServicios.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "frmAF_ComServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1

Set Me.Icon = frmContenedor.Icon
Call sbToolBarIconos(tlb, False)

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False


strSQL = "select cod_servicio,descripcion,estado from afi_comservicios" _
      & " order by cod_servicio"
Call sbCargaGrid(vGrid, 3, strSQL)


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
       vGrid.Row = vGrid.ActiveRow
       vGrid.Col = 1
       strSQL = "delete afi_comservicios where cod_servicio = '" & vGrid.Text & "'"
       glogon.Conection.Execute strSQL
   
       Call Bitacora("Elimina", "Mantenimieto Servicio/Comites : " & vGrid.Text)
   
       Call Form_Load
     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
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


End Sub
Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from afi_comservicios " _
       & " where cod_servicio = '" & vGrid.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into afi_comservicios(cod_servicio,descripcion,estado) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ")"
  
  glogon.Conection.Execute strSQL

  vGrid.Col = 1
  Call Bitacora("Registra", "Mantenimieto Servicio/Comites : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update afi_comservicios set descripcion = '" & vGrid.Text & "',"
 vGrid.Col = 3
 strSQL = strSQL & "estado = " & vGrid.Value & "  where cod_servicio = '" & ""
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 glogon.Conection.Execute strSQL

 vGrid.Col = 1
 Call Bitacora("Modifica", "Mantenimiento Servio/Comites : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function


