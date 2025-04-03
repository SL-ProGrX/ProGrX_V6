VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_AutorizacionTranferencias 
   Caption         =   "Autorización de Tranferencias"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13560
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
   ScaleHeight     =   7545
   ScaleWidth      =   13560
   Begin VB.CheckBox chkRevisados 
      Caption         =   "Revisados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame fraOperaciones 
      Caption         =   "Operaciones"
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   13335
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4812
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   12972
         _Version        =   524288
         _ExtentX        =   22881
         _ExtentY        =   8488
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
         SpreadDesigner  =   "frmCR_AutorizacionTranferencias.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   240
         TabIndex        =   8
         Top             =   5520
         Width           =   1425
         _ExtentX        =   2514
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
   End
   Begin VB.ComboBox cboEtiquetas 
      Height          =   330
      ItemData        =   "frmCR_AutorizacionTranferencias.frx":0F37
      Left            =   2160
      List            =   "frmCR_AutorizacionTranferencias.frx":0F39
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   12000
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
            Picture         =   "frmCR_AutorizacionTranferencias.frx":0F3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12000
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
            Picture         =   "frmCR_AutorizacionTranferencias.frx":779D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFInicio 
      Height          =   330
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Fecha Inicio Búsqueda"
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   244187139
      CurrentDate     =   40361
   End
   Begin MSComctlLib.Toolbar tlbBuscar 
      Height          =   312
      Left            =   7800
      TabIndex        =   9
      Top             =   240
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
   Begin VB.Label Label3 
      Caption         =   "Formalización:"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmCR_AutorizacionTranferencias.frx":DFFF
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Etiqueta"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_AutorizacionTranferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCargaGrid As Boolean
Private mTagAutoriza As String

Private Sub sbParametrosTagAutorizacion()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    strSQL = "select isnull(valor,'') from CRD_PARAMETROS where cod_parametro = '31'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagAutoriza = rs.Fields(0)
    Else
        mTagAutoriza = Empty
    End If
    rs.Close
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
    vModulo = 3
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = vbChecked Then
        Call sbGridMarcarTodo(True)
    Else
        Call sbGridMarcarTodo(False)
    End If
End Sub

Private Sub Form_Load()
    vModulo = 3

    Call sbParametrosTagAutorizacion

    mCargaGrid = True
    
    vGrid.AppearanceStyle = fxGridStyle
    
    vGrid.MaxRows = 0
    vGrid.MaxCols = 11
    
    Call sbCargarCombos
    
    Call Formularios(Me)
    Call RefrescaTags(Me)
       
    Me.Width = 13800
       
End Sub

Private Sub sbCargarCombos()
Dim strSQL As String, rs As New ADODB.Recordset, strPrimero As String

On Error GoTo vError

    strSQL = "SELECT CT.TAG_CODIGO as llave,CT.DESCRIPCION as describe FROM CRD_TAGS CT INNER JOIN CRD_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
           & " INNER JOIN CRD_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
           & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
           & "' order by CT.TAG_CODIGO"
    Call OpenRecordSet(rs, strSQL)

    If Not rs.EOF And Not rs.BOF Then strPrimero = Trim(rs!llave) & " - " & Trim(rs!describe)

    Do While Not rs.EOF
      cboEtiquetas.AddItem Trim(rs!llave) & " - " & Trim(rs!describe)
      rs.MoveNext
    Loop
    If strPrimero <> "" Then cboEtiquetas.Text = strPrimero
    rs.Close
    
    dtpFInicio.Value = fxFechaServidor

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
    
    mCargaGrid = True
    
    strSQL = "SELECT RC.ID_SOLICITUD,RC.FECHAFORP,S.NOMBRE,RC.CODIGO,RC.MONTOSOL,RC.CUOTA,RC.PLAZO,RC.INT," _
            & "case RC.ESTADOSOL when 'R' then 'Recibido' when 'P' then 'Pendiente' else RC.ESTADOSOL end, RC.FECHASOL " _
            & "FROM REG_CREDITOS RC INNER JOIN SOCIOS S ON RC.CEDULA = S.CEDULA " _
            & "WHERE RC.ESTADOSOL = 'F' and ISNULL(RC.AUTORIZA_TRANSFERENCIA,0) = 0 and RC.FECHAFORP = '" & Format(dtpFInicio, "yyyy/mm/dd") _
              
    strSQL = strSQL & "' and dbo.fxCRDValidaTag('" & SIFGlobal.fxCodText(cboEtiquetas) & "',RC.ID_SOLICITUD) > 0"
    
    If chkRevisados.Value = vbChecked Then
        strSQL = strSQL & " and isnull(ANALISTAS_REVISION,0) = 1"
    End If
    
    strSQL = strSQL & " order by RC.ID_SOLICITUD"
        
    Call sbCargaGridCheckIni(vGrid, 10, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    mCargaGrid = False
    chkTodos.Value = Unchecked
    Me.MousePointer = vbDefault
    Exit Sub
error:
    mCargaGrid = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Public Sub sbCargaGridCheckIni(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
'Procedimiento para cargar grids con el check en la primera columna
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

    vGrid.MaxCols = vGridMaxCol + 1
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

Private Sub Form_Resize()
'' Procedimiento para posicionar los controles al max y minimizar la pantalla
On Error GoTo vError
        fraOperaciones.Width = Me.Width - 600
        vGrid.Width = fraOperaciones.Width - 450
        
        fraOperaciones.Height = Me.Height - 1850
        vGrid.Height = fraOperaciones.Height - 1500
        
        tlbAplicar.top = vGrid.top + vGrid.Height + 150
        
    
    Exit Sub

vError:

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If vGrid.MaxRows = 0 Then
        Exit Sub
    End If
    
    If mTagAutoriza = Empty Then
        MsgBox "No está definido en parámetros de créditos la etiqueta de autorización de transferencia"
        Exit Sub
    End If
    
    If MsgBox("Está seguro que sea autorizar la transferencias de las operaciones seleccionadas", vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Call sbAprobarTransferencias
    Call sbCargarListaSolicitudes
    Me.MousePointer = vbDefault
End Sub

Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
If cboEtiquetas.ListCount <= 0 Then Exit Sub

Call sbCargarListaSolicitudes

End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim frm As Form
    If Not mCargaGrid Then
    
        Dim IdSolicitud As Long
        
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
                    Call frm.sbConsultaExterna(IdSolicitud)
                    Exit For
                End If
            Next frm
        Case 3
            Operacion.Operacion = IdSolicitud
            If Operacion.Operacion > 0 Then
                Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)
            End If
        End Select
    
    End If
End Sub

Public Sub sbGridMarcarTodo(Habilita As Boolean)
' Procedimiento para marcar todos los checks en un grid
Dim i As Long

On Error GoTo vError

    vGrid.Row = 1
    vGrid.col = 1
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        vGrid.col = 1
        If Habilita Then
            vGrid.Value = vbChecked
        Else
            vGrid.Value = Unchecked
        End If
    Next i
    Exit Sub
vError:
        MsgBox fxSys_Error_Handler(Err.Description)
    
End Sub

Private Sub sbAprobarTransferencias()
'' Procedimiento para autorizar la tranferencia de la operación

Dim IdSolicitud As Long, Linea As String, i As Integer

On Error GoTo vError

    vGrid.Row = 1
    vGrid.col = 1
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        
        If vGrid.Value = vbChecked Then
        
                vGrid.col = 4
                IdSolicitud = vGrid.Value
                
                If IdSolicitud <> Empty Then
                
                    vGrid.col = 7
                    Linea = vGrid.Value
                    
''                  Se pasa al sp al insertar el tag
'                    Call sbRegCreditosAutorizaTransferencia(IdSolicitud)
                    
                    Call sbCrdOperacionTags(IdSolicitud, Linea, mTagAutoriza, "", "Autorización de Transferencia")
                    
                    Call Bitacora("Autoriza", "Autoriza Transferencia Crédito:" & IdSolicitud)
                    
                End If
                
                vGrid.col = 1
        End If
    Next i
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbRegCreditosAutorizaTransferencia(ByVal Operacion As String)
'' Procedimiento para cambiar en REG_CREDITOS el campo AUTORIZA_TRANSFERENCIA A 1
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    If Operacion = Empty Then
        Exit Sub
    End If
    
    strSQL = "update REG_CREDITOS SET AUTORIZA_TRANSFERENCIA = '1' WHERE ID_SOLICITUD = " & Operacion
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

