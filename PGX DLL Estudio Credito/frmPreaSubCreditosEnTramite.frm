VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmPreaSubCreditosEnTramite 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGridAuto 
      Height          =   2172
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   7452
      _Version        =   524288
      _ExtentX        =   13145
      _ExtentY        =   3831
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
      SpreadDesigner  =   "frmPreaSubCreditosEnTramite.frx":0000
      AppearanceStyle =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2172
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   7452
      _Version        =   524288
      _ExtentX        =   13145
      _ExtentY        =   3831
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
      SpreadDesigner  =   "frmPreaSubCreditosEnTramite.frx":0585
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Totales ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Casos Manuales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Casos Automáticos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   9240
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Créditos ?"
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
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6612
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaSubCreditosEnTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String

Me.Caption = "Expediente : " & gPreAnalisis.Expediente

vGrid.AppearanceStyle = AppearanceStyleVisualStyles
vGridAuto.AppearanceStyle = AppearanceStyleVisualStyles

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

If gPreAnalisis.Tag1 = "C" Then
   lblTipo.Caption = "Cuotas de Créditos Cancelados al Cobro"
Else
   lblTipo.Caption = "Cuotas de Créditos x Cobrar"
End If

strSQL = "select id_solicitud,detalle,cuota from CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO" _
       & " where tipo = 'A' and estado = '" & gPreAnalisis.Tag1 _
       & "' and cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
        
Call sbCargaGrid(vGridAuto, 3, strSQL)
'vGridAuto.MaxRows = vGridAuto.MaxRows - 1

strSQL = "select id_solicitud,detalle,cuota from CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO" _
       & " where tipo = 'M' and estado = '" & gPreAnalisis.Tag1 _
       & "' and cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
Call sbCargaGrid(vGrid, 3, strSQL)

Call sbCalculaTotales

End Sub


Private Sub sbCalculaTotales()
Dim i As Integer, curCuota As Currency

curCuota = 0

For i = 1 To vGridAuto.MaxRows
    vGridAuto.Row = i
    vGridAuto.Col = 3 'Cuota
    curCuota = curCuota + IIf((vGridAuto.Text = ""), 0, vGridAuto.Text)
Next i

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 3 'Cuota
    curCuota = curCuota + IIf((vGrid.Text = ""), 0, vGrid.Text)
Next i


txtCuota.Text = Format(curCuota, "Standard")

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


If vGrid.Text = "" Then  'Insertar

    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
        Exit Function
    End If
    
  strSQL = "select isnull(Max(Id_solicitud),0)+1 as Ultimo from CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO" _
         & " where Tipo = 'M' and cod_preAnalisis = '" _
         & gPreAnalisis.Expediente & "'"
  Call OpenRecordSet(rs, strSQL)
      vGrid.Text = rs!Ultimo
  rs.Close
    
  strSQL = "insert into CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO(cod_PreAnalisis,id_solicitud,tipo,estado,detalle,cuota)" _
         & " values('" & gPreAnalisis.Expediente & "'," & vGrid.Text & ",'M','" & gPreAnalisis.Tag1 & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ")"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Estudio de Credito Cuota en Transito E:" & gPreAnalisis.Expediente & " ID : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO set detalle = '" & vGrid.Text & "',cuota = "
 vGrid.Col = 3
 strSQL = strSQL & CCur(vGrid.Text) & " where cod_PreAnalisis = '" & gPreAnalisis.Expediente _
        & "' and Tipo = 'M' and id_solicitud = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Estudio de Credito Cuota en Transito E:" & gPreAnalisis.Expediente & " ID : " & vGrid.Text)

End If

Call sbCalculaTotales

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
   vGrid.Col = 1
   
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
        Exit Sub
    End If
   
   strSQL = "delete CRD_PREA_DETALLE_CUOTAS_EN_TRANSITO where cod_PreAnalisis = '" & gPreAnalisis.Expediente _
          & "' and Tipo = 'M' and Id_solicitud = " & vGrid.Text
   Call ConectionExecute(strSQL)
   strSQL = vGrid.Text
   vGrid.Col = 1
   Call Bitacora("Elimina", "Estudio de Credito Cuota en Transito E:" & gPreAnalisis.Expediente & " ID : " & vGrid.Text)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1

   Call sbCalculaTotales

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Unload(Cancel As Integer)
GLOBALES.gTag = txtCuota.Text
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

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

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbBorrar
End If

End Sub

Private Sub vGridAutoAuto_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

Dim i As Integer

If vGridAuto.ActiveCol = vGridAuto.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
'  i = fxGuardar
End If

End Sub
