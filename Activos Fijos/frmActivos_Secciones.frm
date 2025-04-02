VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmActivos_Secciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Secciones"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   5295
   End
   Begin FPSpread.vaSpread vGrid 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      _Version        =   393216
      _ExtentX        =   11880
      _ExtentY        =   8493
      _StockProps     =   64
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   2
      SpreadDesigner  =   "frmActivos_Secciones.frx":0000
      VisibleCols     =   500
      VisibleRows     =   500
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisDep"
                  Text            =   "Listado de Departamentos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LisLineas"
                  Text            =   "Líneas x Departamento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmActivos_Secciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vSeccion As String, vPaso As Boolean

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Not vPaso Then Exit Sub

strSQL = "select cod_seccion,descripcion from Activos_secciones" _
      & " where cod_departamento = '" & SIFGlobal.fxSIFCodText(cbo.Text) _
      & "' order by cod_seccion"
Call sbCargaGrid(vGrid, 2, strSQL)

End Sub

Private Sub Form_Activate()
vModulo = 36
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vPaso = False

vModulo = 36
Call Formularios(Me)

strSQL = "select rtrim(cod_departamento) + ' - ' + rtrim(descripcion) as Departamento from Activos_departamentos order by cod_departamento"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cbo.AddItem rs!departamento
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cbo.Text = rs!departamento
End If
rs.Close

Call sbToolBarIconos(tlb)

vPaso = True

Call cbo_Click

Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select coalesce(count(*),0) as Existe from Activos_secciones" _
       & " where cod_seccion = '" & vGrid.Text & "' and cod_departamento = '" _
       & SIFGlobal(cbo.Text) & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into Activos_secciones(cod_departamento,cod_seccion,descripcion) values('" _
         & SIFGlobal(cbo.Text) & "','" & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"

  glogon.Conection.Execute strSQL

  vGrid.Col = 1
'  Call sbBitacora("Registra", "Departamento : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update Activos_secciones set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_departamento = '" & SIFGlobal(cbo.Text) & "' and cod_seccion = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 glogon.Conection.Execute strSQL

  vGrid.Col = 1
'  Call sbBitacora("Modifica", "Departamento : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

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
        vGrid.Col = 1
        strSQL = "delete Activos_secciones where cod_departamento = '" & SIFGlobal(cbo.Text) _
               & "' and cod_seccion = '" & vGrid.Text & "'"
        glogon.Conection.Execute strSQL
        strSQL = vGrid.Text
        vGrid.Col = 2
      '  Call sbBitacora("Elimina", "Departamento : " & strSQL & " - " & vGrid.Text)
        
        vGrid.Col = 1
        Call cbo_Click

     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
'  Case "LisDep"
'     Call sbReportesInv("Departamentos", "Departamentos", "Listado", "")
'  Case "LisLineas"
'     Call sbReportesInv("DeptLineas", "Líneas x Departamentos", "Listado", "")
End Select
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

End Sub






