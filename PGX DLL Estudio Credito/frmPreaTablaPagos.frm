VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaTablaPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tabla de Pagos"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   240
      Top             =   1560
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4095
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   8655
      _Version        =   524288
      _ExtentX        =   15266
      _ExtentY        =   7223
      _StockProps     =   64
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaTablaPagos.frx":0000
      VScrollSpecial  =   -1  'True
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11880
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   18
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Fechas de Pagos de Salario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaTablaPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub cbo_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

strSQL = "select idx,fecha,usuario,inicio,corte,npagos from crd_prea_tabla_pagos" _
       & " where cod_institucion = " & cbo.ItemData(cbo.ListIndex) & " order by inicio desc"
Call sbCargaGrid(vGrid, 6, strSQL)

vError:


End Sub

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()

vModulo = 3 'Modulo de Credito

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Or vGrid.Text = "0" Then
   vGrid.Col = 4
   strSQL = "insert into Crd_Prea_Tabla_pagos(cod_institucion,fecha,usuario,inicio,corte,npagos)" _
          & " values(" & cbo.ItemData(cbo.ListIndex) & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" _
          & Format(vGrid.Text, "yyyy/mm/dd") & "','"
   vGrid.Col = 5
   strSQL = strSQL & Format(vGrid.Text, "yyyy/mm/dd") & "',"
   vGrid.Col = 6
   strSQL = strSQL & vGrid.Text & ")"
   
   Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(max(IDx),0) as ultimo from Crd_Prea_Tabla_pagos"
    Call OpenRecordSet(rs, strSQL)
      vGrid.Col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
    vGrid.Col = 1
    Call Bitacora("Registra", "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text)
    
    MsgBox "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text & ", Registrado satisfactoriamente!", vbInformation
   
      vGrid.Col = 2
      vGrid.Text = fxFechaServidor
      vGrid.Col = 3
      vGrid.Text = glogon.Usuario
   
   Else 'Actualizar
    vGrid.Col = 4
    strSQL = "update Crd_Prea_Tabla_pagos set inicio = '" & Format(vGrid.Text, "yyyy/mm/dd")
    vGrid.Col = 5
    strSQL = strSQL & "',corte = '" & Format(vGrid.Text, "yyyy/mm/dd")
    vGrid.Col = 6
    strSQL = strSQL & "',npagos = " & vGrid.Text & ", Modifica_usuario = '" & glogon.Usuario & "', Modifica_fecha = dbo.MyGetdate()"
    vGrid.Col = 1
    strSQL = strSQL & " where Idx = " & vGrid.Text
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text)
    
    MsgBox "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text & ", modificado satisfactoriamente!", vbInformation
    
      vGrid.Col = 2
      vGrid.Text = fxFechaServidor
      vGrid.Col = 3
      vGrid.Text = glogon.Usuario
    
    
   End If

   vGrid.Col = 1
   fxGuardar = vGrid.Text
   
   Exit Function
   
vError:
 fxGuardar = 0
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Function

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


Dim strSQL As String

vPaso = True

cbo.Clear

strSQL = "select cod_institucion as 'IdX',descripcion as 'ItmX'" _
      & " from instituciones Order by cod_institucion"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = False

Call cbo_Click


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Col = 1
  If vGrid.Text <> "" Then
  
    i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
    If i = vbYes Then
          strSQL = "delete Crd_Prea_Tabla_pagos where IDx = " & vGrid.Text
          Call ConectionExecute(strSQL)
        
          Call Bitacora("Elimina", "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text)
          
          MsgBox "Estudio Credito Tabla de Pago [ID]: " & vGrid.Text & ", Eliminado!", vbInformation
        
          vGrid.DeleteRows vGrid.ActiveRow, 1
          vGrid.MaxRows = vGrid.MaxRows - 1
          If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
    End If
  End If
  
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


