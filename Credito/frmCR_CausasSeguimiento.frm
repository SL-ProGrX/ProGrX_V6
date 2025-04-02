VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_CausasSeguimiento 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Causas de Seguimiento de Tramites"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   7410
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7410
      _ExtentX        =   13070
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
      Height          =   5052
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7212
      _Version        =   524288
      _ExtentX        =   12721
      _ExtentY        =   8911
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_CausasSeguimiento.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   3840
      TabIndex        =   3
      Top             =   480
      Width           =   3252
      _Version        =   1441792
      _ExtentX        =   5741
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Gestión - Solicitudes en tramite:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmCR_CausasSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

If Not vPaso Then Exit Sub

strSQL = "select cod_causas,descripcion,estado from OPERACION_CAUSAS" _
       & " where tipo = '" & Mid(cbo.Text, 1, 1) & "'" _
       & " order by cod_causas"
Call sbCargaGrid(vGrid, 3, strSQL)
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle
 
Call sbToolBarIconos(tlb)

vPaso = False
cbo.AddItem "Pendientes"
cbo.AddItem "Denegadas"
cbo.Text = "Pendientes"
vPaso = True

Call Formularios(Me)
Call RefrescaTags(Me)

Call cbo_Click

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

strSQL = "select isnull(count(*),0) as Existe from OPERACION_CAUSAS " _
       & " where cod_causas = '" & vGrid.Text & "' and tipo = '" & Mid(cbo.Text, 1, 1) & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into OPERACION_CAUSAS(tipo,cod_causas,descripcion,estado) values('" & Mid(cbo.Text, 1, 1) & "','" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ")"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Causa de Seguim.Tramite Cod : " & Trim(vGrid.Text) & " Tipo: " & Mid(cbo.Text, 1, 1))

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update OPERACION_CAUSAS set descripcion = '" & vGrid.Text & "',estado = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where cod_causas = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "' and tipo = '" & Mid(cbo.Text, 1, 1) & "'"
 Call ConectionExecute(strSQL)

  Call Bitacora("Modifica", "Causa de Seguim.Tramite Cod : " & Trim(vGrid.Text) & " Tipo: " & Mid(cbo.Text, 1, 1))

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
        strSQL = "delete OPERACION_CAUSAS where cod_causas = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Causas Seguimiento Tramite : " & strSQL & " - " & vGrid.Text)
        
        vGrid.col = 1
        
        Call cbo_Click
        
     End If
  
  Case "REPORTES"

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





