VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPosEmisores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisores"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   7710
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
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
      Height          =   5652
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7572
      _Version        =   524288
      _ExtentX        =   13356
      _ExtentY        =   9970
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosEmisores.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmPosEmisores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String

Set Me.Icon = frmContenedor.Icon

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb)

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select cod_emisor,descripcion,bloqueo_ck from pv_emisores" _
      & " order by cod_emisor"
Call sbCargaGrid(vGrid, 3, strSQL)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from pv_emisores " _
       & " where cod_emisor = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into pv_emisores(cod_emisor,descripcion,bloqueo_ck) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.col = 3
  strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Entidad Emisora Cod : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update pv_emisores set descripcion = '" & vGrid.Text & "',bloqueo_ck = "
 vGrid.col = 3
 strSQL = strSQL & IIf((vGrid.Text = "1"), 1, 0) & " where cod_emisor = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Entidad Emisora Cod : " & vGrid.Text)

End If
rs.Close


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = 6 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete pv_emisores where cod_emisor = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Tipo Precio: " & strSQL & " - " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select cod_emisor,descripcion,bloqueo_ck from pv_emisores " _
              & " order by cod_emisor"
        Call sbCargaGrid(vGrid, 3, strSQL)

     End If
  
  Case "REPORTES"
    ' Call sbInvReportes("TiposPrecios", "Tipos de Precios", "Listado", "")

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
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





