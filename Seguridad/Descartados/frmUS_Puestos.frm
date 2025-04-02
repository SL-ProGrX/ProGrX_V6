VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUS_Puestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición de Puestos"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
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
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _Version        =   524288
      _ExtentX        =   10610
      _ExtentY        =   8070
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
      MaxCols         =   486
      ScrollBars      =   2
      SpreadDesigner  =   "frmUS_Puestos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmUS_Puestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String
'vModulo = 13

Set Me.Icon = frmContenedor.Icon
vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb)
If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

strSQL = "select cod_puesto,descripcion from us_puestos order by descripcion"
Call sbCargaGrid(vGrid, 2, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Then

    vGrid.Col = 2
    strSQL = "insert into us_puestos(descripcion) values('" & vGrid.Text & "')"
    Call ConectionExecute(strSQL)
  
    vGrid.Col = 2
    Call Bitacora("Registra", "Puesto : " & vGrid.Text)
  
    strSQL = "select max(cod_puesto) as Ultimo from us_puestos"
    Call OpenRecordSet(rs, strSQL)
    fxGuardar = rs!Ultimo
    rs.Close
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update us_puestos set descripcion = '" & UCase(vGrid.Text) & "'"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "' where cod_puesto = " & vGrid.Text
    Call ConectionExecute(strSQL)
    
End If

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
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete us_puestos where cod_puesto = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        vGrid.Col = 2
        Call Bitacora("Elimina", "Puesto : " & vGrid.Text)
        
        vGrid.Col = 1
        
        strSQL = "select cod_puesto,descripcion from us_puestos order by descripcion"
        Call sbCargaGrid(vGrid, 2, strSQL)
     End If

  
  Case "REPORTES"
    ' Call sbReportes("Parametros", Me)
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

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

End Sub







