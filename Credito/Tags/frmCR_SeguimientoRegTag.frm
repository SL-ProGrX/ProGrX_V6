VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_SeguimientoRegTag 
   Caption         =   "Seguimiento de Etiquetas"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   12120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRegTag.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoRegTag.frx":6862
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Operaciones"
      Height          =   7095
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   12255
      Begin VB.Frame fraFiltros 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   8175
         Begin MSComctlLib.Toolbar tlbBuscar 
            Height          =   312
            Left            =   6960
            TabIndex        =   9
            Top             =   120
            Width           =   1188
            _ExtentX        =   2090
            _ExtentY        =   556
            ButtonWidth     =   1640
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Buscar"
                  Key             =   "Buscar"
                  Object.ToolTipText     =   "Buscar"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   330
            Left            =   4080
            TabIndex        =   12
            Top             =   120
            Width           =   1815
            _Version        =   1441793
            _ExtentX        =   3201
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
         Begin XtremeSuiteControls.DateTimePicker dtpFInicio 
            Height          =   330
            Left            =   1440
            TabIndex        =   13
            Top             =   120
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   582
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFFin 
            Height          =   330
            Left            =   2760
            TabIndex        =   14
            Top             =   120
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2355
            _ExtentY        =   582
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin VB.Label Label5 
            Caption         =   "Filtros   >>>"
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   11775
         _Version        =   524288
         _ExtentX        =   20770
         _ExtentY        =   8916
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
         MaxCols         =   493
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_SeguimientoRegTag.frx":D0C4
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   10320
         TabIndex        =   7
         Top             =   6120
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1005
         ButtonWidth     =   2117
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Etiqueta"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtNota 
         Height          =   855
         Left            =   1800
         TabIndex        =   15
         Top             =   6120
         Width           =   8415
         _Version        =   1441793
         _ExtentX        =   14843
         _ExtentY        =   1508
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblObservacion 
         Caption         =   "Observación"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   6120
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.ComboBox cboEtiquetas 
      Height          =   330
      Left            =   2520
      TabIndex        =   11
      Top             =   720
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
   Begin VB.Image Image1 
      Height          =   645
      Left            =   240
      Picture         =   "frmCR_SeguimientoRegTag.frx":E0DA
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "Etiqueta"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5640
      TabIndex        =   2
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label lblUsuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_SeguimientoRegTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean



Private Sub cboEtiquetas_Click()

If vPaso Then Exit Sub
    
    vGrid.MaxRows = 0
    Call sbCargarListaSolicitudes

End Sub

Private Sub Form_Activate()

vModulo = 8
End Sub

Private Sub Form_Load()

vModulo = 8

    vPaso = True
    
    lblUsuario.Caption = glogon.Usuario
    lblNombre.Caption = fxNombreUsuario
    
    vGrid.AppearanceStyle = fxGridStyle
    
    vGrid.MaxRows = 0
    vGrid.MaxCols = 14
    
    Call sbCargarCombos
       
End Sub

Private Sub Form_Resize()
'' Procedimiento para posicionar los controles al max y minimizar la pantalla
On Error GoTo vError
    
    If Me.Width > 10000 Then
    
        fraOperaciones.Width = Me.Width - 500
        fraOperaciones.Height = Me.Height - 1900
        
        vGrid.Width = fraOperaciones.Width - 400
        
        fraFiltros.Left = fraOperaciones.Width - fraFiltros.Width - 250
        
        vGrid.Height = fraOperaciones.Height - 2000
        
        lblObservacion.top = vGrid.top + vGrid.Height + 300
        txtNota.top = lblObservacion.top
        tlbAplicar.top = lblObservacion.top
        
        txtNota.Width = vGrid.Width - 3000
        
        tlbAplicar.Left = txtNota.Left + txtNota.Width + 300
        

    End If
    
    Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxNombreUsuario() As String
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select DESCRIPCION from USUARIOS where NOMBRE = '" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxNombreUsuario = rs.Fields(0)
    Else
        fxNombreUsuario = Empty
    End If

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


Private Sub sbCargarCombos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strPrimero As String

On Error GoTo vError

    vPaso = True
    
    strSQL = "SELECT CT.TAG_CODIGO as llave,CT.DESCRIPCION as describe FROM CRD_TAGS CT INNER JOIN CRD_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
           & " INNER JOIN CRD_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
           & " WHERE CT.ACTIVO = 1 AND NOT ISNULL(CT.COD_REQUISITO,'') = '' AND CGU.USUARIO = '" & glogon.Usuario _
           & "' order by CT.TAG_CODIGO"
    Call OpenRecordSet(rs, strSQL)

    If Not rs.EOF And Not rs.BOF Then strPrimero = Trim(rs!llave) & " - " & Trim(rs!describe)

    Do While Not rs.EOF
      cboEtiquetas.AddItem Trim(rs!llave) & " - " & Trim(rs!describe)
      rs.MoveNext
    Loop
    If strPrimero <> "" Then cboEtiquetas.Text = strPrimero
    rs.Close
    
    cboEstado.Clear
    cboEstado.AddItem ("Todos")
    cboEstado.AddItem ("Recibida")
    cboEstado.AddItem ("Pendiente")
    cboEstado.Text = "Todos"
    
    dtpFInicio.Value = fxFechaServidor
    dtpFFin.Value = dtpFInicio.Value

    vPaso = False
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarListaSolicitudes()
' Carga Lista de operaciones
    Dim strSQL As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    
    Me.MousePointer = vbHourglass
    vPaso = True
    
    strSQL = "SELECT RC.ID_SOLICITUD,RC.FECHASOL,S.NOMBRE,RC.CODIGO,RC.MONTOSOL,RC.CUOTA,RC.PLAZO,RC.INT," _
            & " case RC.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else RC.ESTADOSOL end, RC.FECHASOL,Ofi.DESCRIPCION" _
            & " FROM REG_CREDITOS RC INNER JOIN SOCIOS S ON RC.CEDULA = S.CEDULA " _
            & " INNER JOIN OPERACION_REQUISITOS ORE ON RC.ID_SOLICITUD = ORE.ID_SOLICITUD " _
            & " INNER JOIN REQUISITOS_ADICIONALES REQ ON ORE.COD_REQUISITO = REQ.COD_REQUISITO " _
            & " INNER JOIN SIF_OFICINAS Ofi ON RC.COD_OFICINA_R = Ofi.COD_OFICINA " _
            & " WHERE ORE.estado = 0 AND RC.FECHASOL BETWEEN '" & Format(dtpFInicio, "yyyy/mm/dd") & "' AND '" _
            & Format(dtpFFin, "yyyy/mm/dd") & "' AND ORE.COD_REQUISITO = '" & fxRequisitoTag & "'"
            
    Select Case cboEstado.Text
    Case "Todos"
        strSQL = strSQL & " and RC.ESTADOSOL in ('P','R')"
    Case "Recibida"
        strSQL = strSQL & " and RC.ESTADOSOL = 'R'"
    Case "Pendiente"
        strSQL = strSQL & " and RC.ESTADOSOL = 'P'"
    End Select
        
    Call sbCargaGridCheckIni(vGrid, 12, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    vPaso = False
    
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub

Private Function fxRequisitoTag() As String
Dim strSQL As String, rs As New ADODB.Recordset
On Error GoTo vError

    strSQL = "select COD_REQUISITO from CRD_TAGS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        fxRequisitoTag = rs.Fields(0)
    Else
        fxRequisitoTag = Empty
    End If

Exit Function
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function


Public Sub sbCargaGridCheckIni(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
'Procedimiento para cargar grids con el check en la primera columna
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

'    vGrid.MaxCols = vGridMaxCol + 1
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.col = i
     vGrid.Text = ""
    Next i
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      vGrid.Row = vGrid.MaxRows
      For i = 4 To vGrid.MaxCols
        vGrid.col = i
        vGrid.Text = CStr(rs.Fields(i - 4).Value & "")
      Next i
      vGrid.MaxRows = vGrid.MaxRows + 1
      rs.MoveNext
    Loop
    rs.Close
    Exit Sub

vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If MsgBox("Está seguro que sea aplicar la etiqueta en las operaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Call sbIncluirEtiquetas
    Call sbCargarListaSolicitudes
End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call sbCargarListaSolicitudes
End Sub

Private Sub sbIncluirEtiquetas()
'' Procedimiento para Ingresar las etiquetas marcadas en las operaciones seleccionadas

Dim IdSolicitud As Long, Linea As String, vEtiqueta As String, i As Long

On Error GoTo vError

vEtiqueta = SIFGlobal.fxCodText(cboEtiquetas.Text)

    vGrid.Row = 1
    vGrid.col = 1
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        
        If vGrid.Value = 1 Then
        
                vGrid.col = 4
                IdSolicitud = vGrid.Value
                
                If IdSolicitud <> Empty Then
                
                    vGrid.col = 7
                    Linea = vGrid.Value
                    
                    Call sbCrdOperacionTags(IdSolicitud, Linea, vEtiqueta, "", txtNota.Text)
                    
                End If
                
                vGrid.col = 1
        End If
    Next i
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim IdSolicitud As Long, frm As Form
    
    
If vPaso Then Exit Sub

    vGrid.Row = Row
    vGrid.col = 4
    
    If vGrid.Value = Empty Then Exit Sub
    
    IdSolicitud = vGrid.Value
    
    If IdSolicitud = Empty Then Exit Sub
    
    Select Case col
    Case 2
        Call sbFormsCall("frmCR_SeguimientoTramites")
        For Each frm In Forms
            If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                Exit For
            End If
        Next frm
        
        Call frm.sbConsultaExterna(IdSolicitud)
    
    
    Case 3
        Operacion.Operacion = IdSolicitud
        If Operacion.Operacion > 0 Then
            Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)
        End If
    End Select
    
End Sub
