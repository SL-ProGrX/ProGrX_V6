VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSeguros_TiposCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Cuentas (Bancos / Deducción Planilla)"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10215
      _Version        =   524288
      _ExtentX        =   18018
      _ExtentY        =   8070
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
      SpreadDesigner  =   "frmSeguros_TiposCuentas.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Cuentas (Deducción o Vía de Cobro)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   8775
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmSeguros_TiposCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 17
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select TIPO_COBRO,descripcion, case when forma_pago = 'PLA' then 'Planilla'" _
       & " when Forma_Pago = 'CAR' then 'Cargo Automático' " _
       & " when Forma_Pago = 'DBC' Then 'Débito en Cuenta' else Forma_Pago end as 'FormaPago' " _
       & ",codigo_deduccion ,Activo" _
       & " from SEGUROS_TIPOS_COBRO" _
       & " order by TIPO_COBRO"
Call sbCargaGrid(vGrid, 5, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxMetodoFormaPago(pDescripcion As String) As String

Select Case pDescripcion
  Case "Planilla"
     pDescripcion = "PLA"
  Case "Cargo Automático"
     pDescripcion = "CAR"
  Case "Débito en Cuenta"
     pDescripcion = "DBC"
  Case Else
     pDescripcion = pDescripcion
End Select

fxMetodoFormaPago = pDescripcion
End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from SEGUROS_TIPOS_COBRO" _
       & " where TIPO_COBRO = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "Insert SEGUROS_TIPOS_COBRO(TIPO_COBRO,descripcion,forma_pago,codigo_deduccion, Activo,Plazo,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & fxMetodoFormaPago(vGrid.Text) & "','"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",0,'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Tipo de Cobro:" & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update SEGUROS_TIPOS_COBRO set descripcion = '" & vGrid.Text & "', Forma_Pago = '"
 vGrid.Col = 3
 strSQL = strSQL & fxMetodoFormaPago(vGrid.Text) & "', Codigo_Deduccion = '"
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & "', Activo = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where TIPO_COBRO = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Tipo de Cobro:" & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete SEGUROS_TIPOS_COBRO where TIPO_COBRO = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Cobro:" & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub









