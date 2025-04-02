VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPosReparacionTipoServicio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SR : Tipos de Servicios"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
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
      Height          =   4932
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7572
      _Version        =   524288
      _ExtentX        =   13356
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosReparacionTipoServicio.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmPosReparacionTipoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle

Set Me.Icon = frmContenedor.Icon
Call sbToolBarIconos(tlb)

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

Call sbCargar

End Sub


Private Sub sbCargar()
Dim strSQL As String

On Error GoTo vError

strSQL = "select Tipo_Servicio,descripcion,activo from POS_REPARACION_TIPOS_SERVICIOS" _
      & " order by Tipo_Servicio"
Call sbCargaGrid(vGrid, 3, strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from POS_REPARACION_TIPOS_SERVICIOS " _
       & " where Tipo_Servicio = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into POS_REPARACION_TIPOS_SERVICIOS(Tipo_Servicio,descripcion,activo) values('" _
         & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "SR. Tipo de Servicio: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update POS_REPARACION_TIPOS_SERVICIOS set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where Tipo_Servicio = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "SR. Tipo de Servicio: " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   strSQL = "delete POS_REPARACION_TIPOS_SERVICIOS where Tipo_Servicio = '" & vGrid.Text & "'"
   Call ConectionExecute(strSQL)
   strSQL = vGrid.Text
   vGrid.col = 1
   Call Bitacora("Elimina", "SR. Tipo de Servicio : " & vGrid.Text)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     
     Call sbBorrar
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

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

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbBorrar
End If

End Sub






