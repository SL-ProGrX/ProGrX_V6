VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCC_PlanillaCtaCorreccion 
   Caption         =   "Planillas: Ajuste de Cuotas"
   ClientHeight    =   8544
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   12588
   LinkTopic       =   "Form1"
   ScaleHeight     =   8544
   ScaleWidth      =   12588
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optBitacora 
      Appearance      =   0  'Flat
      Caption         =   "Bitacora de Cambios"
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
      Left            =   4560
      TabIndex        =   20
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtLineas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Text            =   "100"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.OptionButton optCuotas 
      Appearance      =   0  'Flat
      Caption         =   "Cuotas Enviadas"
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
      Left            =   2400
      TabIndex        =   17
      Top             =   1680
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox txtUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      TabIndex        =   16
      Top             =   2605
      Width           =   1455
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2605
      Width           =   3615
   End
   Begin VB.TextBox txtcedula 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   13
      Top             =   2605
      Width           =   1695
   End
   Begin VB.TextBox txtLinea 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   2605
      Width           =   1245
   End
   Begin VB.TextBox txtOp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   11
      Top             =   2605
      Width           =   1455
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   8055
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.Toolbar tlbAutorizado 
      Height          =   336
      Left            =   9600
      TabIndex        =   22
      Top             =   7800
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   593
      ButtonWidth     =   4297
      ButtonHeight    =   550
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Autorizado para Modificar"
            Key             =   "Autorizado"
            Object.ToolTipText     =   "Informe de Transferencia"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4572
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   12012
      _Version        =   524288
      _ExtentX        =   21188
      _ExtentY        =   8065
      _StockProps     =   64
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      SpreadDesigner  =   "frmCC_PlanillaCtaCorreccion.frx":0000
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio de Cuotas enviadas al Cobro"
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
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   21
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   19
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
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
      Height          =   315
      Index           =   5
      Left            =   4680
      TabIndex        =   14
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
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
      Height          =   315
      Index           =   6
      Left            =   8280
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cédula"
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
      Height          =   315
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Linea"
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
      Height          =   315
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   9840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
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
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Filtros de Consulta: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmCC_PlanillaCtaCorreccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean
Dim vTamanoForm As Double
Dim aColumnas(6) As Double
Dim vHanchoGrid As Double, vAltoGrid As Double

Private Sub cboInstitucion_Click()

If vPaso Then Exit Sub

If optCuotas Then
  Call sbConsulta(1)
Else
  Call sbConsulta(2)
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim vFecha As Currency

On Error GoTo vError

vFecha = txtProceso.Text


If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       vFecha = fxFechaProcesoSiguiente(vFecha)
    Else
       vFecha = fxFechaProcesoAnterior(vFecha)
    End If
    
    txtProceso.Text = vFecha
      
    Call cboInstitucion_Click
End If



vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

vModulo = 10

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vHanchoGrid = vGrid.Width
vAltoGrid = vGrid.Height

vTamanoForm = Me.ScaleWidth

vPaso = True

txtProceso.Text = GLOBALES.glngFechaCR

vGrid.TextTip = TextTipFixed
vGrid.TextTipDelay = 1000


For i = 0 To 6
    aColumnas(i) = vGrid.ColWidth(i + 1)
Next i


 
strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones order by descripcion"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)

strSQL = "select rtrim(descripcion) as Descripcion from instituciones where cod_institucion = " & GLOBALES.gInstitucion
Call OpenRecordSet(rs, strSQL)
    cboInstitucion.Text = rs!Descripcion
rs.Close
vGrid.AppearanceStyle = fxGridStyle

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbConsulta(vTipo As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


vGrid.MaxRows = 0
vGrid.MaxRows = 1

If vTipo = 1 Then
  strSQL = "select  top " & txtLineas.Text & " D.*,isnull(S.nombre,'') as 'Nombre' from PRM_ENVIADO_DETALLE D left join socios S on D.cedula = S.cedula where "
Else
  strSQL = "select  top " & txtLineas.Text & " D.*,isnull(S.nombre,'') as 'Nombre' from prm_cambios D left join socios S on D.cedula = S.cedula where "
End If

If Trim(txtOp.Text) <> "" Then
  strSQL = strSQL & " D.id_solicitud  = '" & txtOp & "' "
End If


If Trim(txtLinea) <> "" Then
  If strSQL <> Empty And strSQL <> "and" And Trim(Mid(strSQL, (Len(strSQL)) - 6, 6)) <> "where" Then strSQL = strSQL & " and "
  'If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
  strSQL = strSQL & " D.codigo  = '" & txtLinea & "' "
End If
  
If Trim(txtcedula) <> "" Then
  If strSQL <> Empty And strSQL <> "and" And Trim(Mid(strSQL, (Len(strSQL)) - 6, 6)) <> "where" Then strSQL = strSQL & " and "
  strSQL = strSQL & " D.cedula  = '" & txtcedula & "' "
End If
  


If Trim(txtNombre) <> "" Then
  If strSQL <> Empty And strSQL <> "and" And Trim(Mid(strSQL, (Len(strSQL)) - 6, 6)) <> "where" Then strSQL = strSQL & " and "
  'If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
  strSQL = strSQL & " D.cedula  = '" & txtcedula & "' "
End If
  
If Trim(txtUsuario) <> "" Then
  If strSQL <> Empty And strSQL <> "and" And Trim(Mid(strSQL, (Len(strSQL)) - 6, 6)) <> "where" Then strSQL = strSQL & " and "
  'If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
  strSQL = strSQL & " D.registro_usuario  = '" & txtUsuario & "' "
End If
  
  
  
   If strSQL <> Empty And strSQL <> "and" And Trim(Mid(strSQL, (Len(strSQL)) - 6, 6)) <> "where" Then strSQL = strSQL & " and "
  'If strSQL <> Empty And strSQL <> "and" Then strSQL = strSQL & " and "
    If vTipo = 1 Then
       strSQL = strSQL & " D.cod_institucion   = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & "  and D.fecpro  = " & txtProceso & ""
    Else
       strSQL = strSQL & " D.cod_institucion   = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & "  and D.proceso  = " & txtProceso & ""
    End If



Me.MousePointer = vbHourglass


Call OpenRecordSet(rs, strSQL)

'ID_CONSECUTIVO , Id_solicitud, CODIGO, FECPRO, CEDULA, CUOTA, MOROSIDAD, COD_INSTITUCION

vGrid.MaxRows = 0
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  vGrid.Col = 1
  
  If vTipo = 1 Then
    vGrid.Text = Trim(rs!ID_CONSECUTIVO)
  Else
    vGrid.Text = Trim(rs!Id_seq)
  End If
  
  vGrid.Col = 2
  If vTipo = 1 Then
     vGrid.Text = Trim(rs!cod_deduccion & "")
  Else
     vGrid.Text = Trim(rs!Proceso)
  End If
 
  
  
  vGrid.Col = 3
  vGrid.Text = Trim(rs!Id_solicitud)
  vGrid.Col = 4
   If vTipo = 1 Then
      vGrid.Text = rs!codigo
  Else
    vGrid.Text = Trim(rs!Linea)
  End If
 
  vGrid.Col = 5
  vGrid.Text = rs!Cedula
  vGrid.Col = 6
  vGrid.Text = rs!Nombre ' fxNombre(rs!Cedula)
  vGrid.Col = 7
  If vTipo = 1 Then
    vGrid.Value = rs!MOROSIDAD
  Else
    vGrid.Text = rs!indicador_mora
  End If
  vGrid.Col = 8
  If vTipo = 1 Then
    vGrid.Text = Format(rs!Cuota, "Standard")
    vGrid.CellTag = Format(rs!Cuota, "Standard")
  Else
    vGrid.Text = Format(rs!Cuota_nueva, "Standard")
    vGrid.CellTag = Format(rs!Cuota_nueva, "Standard")
  End If
  If vTipo = 2 Then
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario " & rs!Registro_Usuario & vbCrLf & "Fecha  " & Format(Mid(rs!Registro_Fecha, 1, 19), "dd/mmmm/yyyy hh:mm ampm") & vbCrLf & "Cuota Ant.  " & rs!Cuota_Anterior
  End If
  
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


Exit Sub


vError:

        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Resize()
Dim i As Integer

On Error GoTo vError

vGrid.Width = Me.ScaleWidth - 300
vGrid.Height = Me.ScaleHeight - 4000
vGrid.ColHeadersAutoText = DispBlank
imgBanner.Width = Me.Width

If vTamanoForm < Me.ScaleWidth Then
    
    txtLineas.Move txtLineas.Left, vGrid.Width - 9280, txtLineas.Width, txtLineas.Height
    lblFecha(7).Move lblFecha(7).Left, vGrid.Width - 9280, lblFecha(7).Width, lblFecha(7).Height
    
    For i = 1 To 7
        vGrid.ColWidth(i) = aColumnas(i - 1) + Me.ScaleWidth / 3920
    Next i
    
ElseIf vTamanoForm = Me.ScaleWidth Then
    
    txtLineas.Move txtLineas.Left, vGrid.Width - 4200, txtLineas.Width, txtLineas.Height
    lblFecha(7).Move lblFecha(7).Left, vGrid.Width - 4200, lblFecha(7).Width, lblFecha(7).Height
    
    For i = 1 To 7
       vGrid.ColWidth(i) = aColumnas(i - 1)
    Next i
    
    vGrid.Height = vAltoGrid
    vGrid.Width = vHanchoGrid
    
    tlbAutorizado.Top = txtLinea.Top
    
End If
Exit Sub

vError:
Resume Next

End Sub

Private Sub optBitacora_Click()

   txtUsuario.Locked = False
   txtUsuario.BackColor = vbWhite
   vGrid.MaxRows = 0
   
   
End Sub

Private Sub optCuotas_Click()
   txtUsuario.Locked = True
   txtUsuario.BackColor = &HC0C0C0
   txtUsuario.Text = Empty
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNombre.SetFocus
End Sub

Private Sub txtCedula_LostFocus()
Call cboInstitucion_Click
End Sub

Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtcedula.SetFocus
End Sub

Private Sub txtLinea_LostFocus()
Call cboInstitucion_Click
End Sub

Private Sub txtLineas_Change()
If Not IsNumeric(txtLineas) Then
   txtLineas.Text = 100
End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If txtUsuario.Locked = False Then txtUsuario.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtcedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
  txtcedula.SetFocus
End If
End Sub






Private Sub txtOp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLinea.SetFocus
End Sub

Private Sub txtOp_LostFocus()
 Call cboInstitucion_Click
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cboInstitucion_Click
End Sub

Private Sub txtUsuario_LostFocus()
Call cboInstitucion_Click
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim strIdConsec As String, vProceso As Long, curVariacion As Currency
Dim vOperacion As Long, iMora As Long, vInstitucion As Long
Dim strCedula As String, strLinea As String, strDeduccionId As String

On Error GoTo vError



vProceso = txtProceso.Text
vInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)

If vGrid.ActiveCol = vGrid.MaxCols And KeyCode = vbKeyReturn _
    And Trim(vGrid.Text) <> "" And optCuotas.Value = True Then
    
    If tlbAutorizado.Buttons.Item(1).Enabled = False Then
       MsgBox "Su usuario no está autorizado para modificar cuotas de planillas...!", vbExclamation
       Exit Sub
    End If
    
    vGrid.Row = vGrid.ActiveRow
    
    vGrid.Col = 1
    strIdConsec = vGrid.Text
    vGrid.Col = 2
    strDeduccionId = vGrid.Text
    vGrid.Col = 3
    vOperacion = vGrid.Text
    vGrid.Col = 4
    strLinea = vGrid.Text
    vGrid.Col = 5
    strCedula = vGrid.Text
    vGrid.Col = 7
    iMora = vGrid.Value
    
    vGrid.Col = vGrid.MaxCols

    If vGrid.CellTag <> vGrid.Text Then
        
        curVariacion = CCur(vGrid.Text) - CCur(vGrid.CellTag)

        If curVariacion <> 0 Then
            strSQL = "exec spPrm_CreditoCambiosManuales_Registro " & vInstitucion & "," & vProceso & ",'" & glogon.Usuario & "'," _
                   & vOperacion & ",'" & strLinea & "','" & strCedula & "'," & CCur(vGrid.Text) & "," & CCur(vGrid.CellTag) _
                   & ",'" & strIdConsec & "'," & iMora & ",'" & strDeduccionId & "'"
        
            Call ConectionExecute(strSQL)
        End If
        
        vGrid.CellTag = vGrid.Text
'        Call cboInstitucion_Click
    End If
        
End If


Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
