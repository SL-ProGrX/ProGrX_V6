VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmPreaEstadoPreanalisis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado del Estudio de Crédito"
   ClientHeight    =   6216
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   9264
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6216
   ScaleWidth      =   9264
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3372
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   8892
      _Version        =   1245187
      _ExtentX        =   15684
      _ExtentY        =   5948
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton cmb_Aceptar 
      Height          =   612
      Left            =   7320
      TabIndex        =   1
      Top             =   5400
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmPreaEstadoPreanalisis.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Recibido"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Pendiente"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   2
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Aprobado"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   372
      Index           =   3
      Left            =   7080
      TabIndex        =   6
      Top             =   1200
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Denegado"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   624
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   12732
      _Version        =   1245187
      _ExtentX        =   22458
      _ExtentY        =   1101
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Resolución"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   432
      Index           =   0
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   3768
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaEstadoPreanalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public m_estadoPreanalisis As String

Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Private mExpediente As String, vId_Comite As Integer
Dim pTipo As String, vPaso As Boolean


Private Sub btnOpcion_Click(Index As Integer)
Dim i As Integer

For i = 0 To btnOpcion.Count - 1
   btnOpcion.Item(i).Checked = False
Next i
btnOpcion.Item(Index).Checked = True

Select Case Index
  Case 0 'Recibido
    pTipo = "R"
  Case 1 'Pendiente
    pTipo = "P"
  Case 2 'Aprobado
    pTipo = "A"
  Case 3 'Denegado
    pTipo = "D"
End Select

Call sbListaCausas

End Sub

Private Sub cmb_Aceptar_Click()
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim vId_Comite As Integer

If m_estadoPreanalisis = "A" Then
    MsgBox "El Análisis seleccionado ya fue aprobado.", vbCritical, gMsgTitulo
    Exit Sub
End If

'Aprobado o Denegado: Valida
If btnOpcion.Item(2).Checked = True Or btnOpcion.Item(3).Checked = True Then
    If Not fxValida Then
        Exit Sub
    End If
End If



'Proceso de Aprobado
If btnOpcion.Item(2).Checked = True Then

    If fxValidaComiteAutorizadores Then
    
        frmPreaAutorizaciones.Show vbModal
    
        If Not fxValidaAutorizadoresMarcados Then
            MsgBox "Debe seleccionar al menos un autorizador para este comité"
            Exit Sub
        End If
        
    End If
    
End If


Call sbGuardar

End Sub

Private Sub sbListaCausas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


lsw.ListItems.Clear

vPaso = True


strSQL = "select Cg.COD_CAUSAS, Cg.DESCRIPCION,case when isnull(Pa.Cod_Causas,'No Existe') = 'No Existe' then 0 else 1 end as 'Check'" _
       & " , Pa.Registro_Fecha, Pa.Registro_Usuario " _
       & " from OPERACION_CAUSAS Cg " _
       & "       left join CRD_PREA_GESTION Pa on Cg.COD_CAUSAS = Pa.COD_CAUSAS and Cg.TIPO = Pa.TIPO" _
       & "             and  Pa.COD_PREANALISIS = '" & mExpediente _
       & "' Where Cg.TIPO = '" & pTipo & "'" _
       & " order by isnull(Pa.REGISTRO_FECHA,getdate()) asc, Cg.Cod_Causas"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Causas)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!registro_Fecha & ""
     itmX.SubItems(3) = rs!registro_usuario & ""
     itmX.Checked = IIf((rs!Check = 1), True, False)
     If itmX.Checked Then itmX.ForeColor = vbBlue
 rs.MoveNext
Loop
rs.Close
 
vPaso = False
 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Me.Caption = "Expediente : " & gPreAnalisis.Expediente

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Código", 1200
lsw.ColumnHeaders.Add , , "Descripción", 3200
lsw.ColumnHeaders.Add , , "Fecha", 2800
lsw.ColumnHeaders.Add , , "Usuario", 2800


If InStr(1, gPreAnalisis.Expediente, "-", vbTextCompare) > 0 Then
    mExpediente = fxDeCodificaPrimaryKey(gPreAnalisis.Expediente, 1, "-")
Else
    mExpediente = gPreAnalisis.Expediente
End If


Select Case m_estadoPreanalisis
  Case "R"
    Call btnOpcion_Click(0)
  Case "P"
    Call btnOpcion_Click(1)
  Case "A"
    Call btnOpcion_Click(2)
  Case "D"
    Call btnOpcion_Click(3)
End Select
    
End Sub

Private Sub sbGuardar()
Dim StrUpdate1 As String
Dim StrUpdate2 As String
Dim StrSet As String
Dim vEstado As String
On Error GoTo vError

StrUpdate1 = "Update CRD_PREA_PREANALISIS SET "
StrUpdate2 = StrUpdate1


Select Case True
  Case btnOpcion.Item(0).Checked  'Recibido
    vEstado = "R"
 Case btnOpcion.Item(1).Checked 'Pendiente
    vEstado = "P"
 Case btnOpcion.Item(2).Checked 'Aprobado
    vEstado = "A"
 Case btnOpcion.Item(3).Checked 'Denegado
    vEstado = "D"
End Select


StrSet = "ESTADO = " & fxFormatearValor(vEstado, Caracter)
StrSet = StrSet & ", USUARIO_GESTION = " & fxFormatearValor(glogon.Usuario, Caracter) & ",  FECHA_GESTION = dbo.MyGetdate()"

StrUpdate1 = StrUpdate1 & StrSet & " where COD_PREANALISIS = " & fxFormatearValor(mExpediente, Caracter)

StrUpdate1 = StrUpdate1 & " or COD_PREANALISIS_REF = " & fxFormatearValor(mExpediente, Caracter)

m_estadoPreanalisis = vEstado

    
If execSql(StrUpdate1, False) Then
    
  '  StrUpdate2 = StrUpdate2 & StrSet & " where COD_PREANALISIS_REF = " & fxFormatearValor(mExpediente, Caracter)
    
    'Call execSql(StrUpdate2, False)
    
'    If vEstado = "A" Or vEstado = "D" Then
        sbInsertarTag (vEstado)
'    End If
    
    MsgBox "La información fue actualizada correctamente.", vbInformation, gMsgTitulo
    UnLoad Me
End If


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub sbInsertarTag(ByVal vEstado As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim Cod_Parametro As String, Tag As String, LineaTag As Integer, NotaTag As String, Cod_Linea As String

    On Error GoTo vError
    
    Select Case vEstado
        Case "A"
            Cod_Parametro = "01"
        Case "D"
            Cod_Parametro = "02"
    End Select
    
    If Not Cod_Parametro = Empty Then
    
        strSQL = "select isnull(valor,'') as valor from CRD_COMITES_PARAMETROS where COD_PARAMETRO ='" & Cod_Parametro & "'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
            Tag = Trim(rs!Valor)
        Else
            MsgBox "No existe en parámetros la información del tag asignado para este movimiento"
            Exit Sub
        End If
        rs.Close
    
        If Tag = Empty Then
            MsgBox "No está definido en parámetros, el tag para este movimiento"
            Exit Sub
        End If
        
        NotaTag = glogon.Usuario & " operación realizada de estudio crediticio"
        
    Else
           
        Tag = "S16"
         
        Select Case vEstado
        Case "P"
            NotaTag = glogon.Usuario & " cambio de estado del estudio crediticio a pendiente"
        Case "R"
            NotaTag = glogon.Usuario & " cambio de estado del estudio crediticio a recibido"
        End Select
        
    End If
    
    strSQL = "select count(*) from crd_tags where tag_codigo = '" & Tag & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.Fields(0) = 0 Then
        MsgBox "El tag definido en parámetros para este movimiento, no existe en el catalogo de tags"
        Exit Sub
    End If
    rs.Close
    
    strSQL = "select isnull(max(linea),0)+1 as Linea from CRD_PREA_TAGS where cod_preanalisis = " & fxFormatearValor(mExpediente, Caracter)
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        LineaTag = rs!Linea
    End If
    rs.Close
    
    strSQL = "select Cod_Linea from CRD_PREA_PREANALISIS where COD_PREANALISIS = " & fxFormatearValor(mExpediente, Caracter)
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        Cod_Linea = rs!Cod_Linea
    End If
    
    ' Insertar el Tag
    strSQL = "insert CRD_PREA_TAGS (LINEA,CODIGO,COD_PREANALISIS,TAG_CODIGO,ASIGNADO_A,REGISTRO_FECHA,REGISTRO_USUARIO,NOTAS)" _
             & "values(" & LineaTag _
             & ",'" & Trim(Cod_Linea) _
             & "'," & fxFormatearValor(mExpediente, Caracter) _
             & ",'" & Tag _
             & "','','" & Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss") _
             & "','" & glogon.Usuario _
             & "','" & NotaTag & "')"
    
    Call ConectionExecute(strSQL)
      
    Exit Sub
vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical
      
End Sub

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

    fxValida = True

    vMensaje = ""

    'Consulta el Comite del preanalisis
    strSQL = "select dbo.fxCrd_Comites_Valida_Resolucion(ID_COMITE, COD_LINEA, GARANTIA, MONTO, '" & glogon.Usuario & "') as 'Mensaje' " _
           & " from CRD_PREA_PREANALISIS" _
           & " where COD_PREANALISIS = '" & mExpediente & "'"
    Call OpenRecordSet(rs, strSQL)
    
    vMensaje = rs!Mensaje
    
    rs.Close


If Len(vMensaje) > 0 Then
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function


Private Function fxValidaComiteAutorizadores() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset

    fxValidaComiteAutorizadores = False

    'Consulta el Comite del preanalisis
    strSQL = "select isnull(Id_comite,0) as Id_Comite from CRD_PREA_PREANALISIS where COD_PREANALISIS = '" & mExpediente & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        vId_Comite = rs!ID_COMITE
    Else
        vId_Comite = 0
    End If
    
    rs.Close
    
    'Verifica si el comite tiene autorizadores
    strSQL = "select COUNT(*) AS 'Cantidad' from CRD_COMITES_AUTORIZADORES where ID_COMITE = " & vId_Comite
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs!Cantidad = 0 Then
            fxValidaComiteAutorizadores = False
        Else
            fxValidaComiteAutorizadores = True
        End If
    End If
    
    rs.Close

End Function

Private Function fxValidaAutorizadoresMarcados() As Boolean
    Dim strSQL As String, rs As New ADODB.Recordset

    fxValidaAutorizadoresMarcados = False
    
    'Verifica si se seleccionaron autorizadores
    strSQL = "select COUNT(*) AS 'Cantidad' from CRD_PREA_AUTORIZADORES where COD_PREANALISIS = '" & mExpediente & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If rs!Cantidad > 0 Then
            fxValidaAutorizadoresMarcados = True
        Else
            fxValidaAutorizadoresMarcados = False
        End If
    End If
    
    rs.Close

End Function

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert CRD_PREA_GESTION(cod_causas,tipo,cod_preanalisis,codigo,registro_fecha,registro_usuario) values('" _
           & Item.Text & "','" & pTipo & "','" & mExpediente _
           & "','" & mCod_linea & "',dbo.Mygetdate(), '" & glogon.Usuario & "')"
Else
  Call Bitacora("Elimina", "Causa SGT: " & Item.Text & ", Expediente: " & mExpediente)
    
  strSQL = "delete CRD_PREA_GESTION where cod_causas = '" & Item.Text & "' and tipo = '" _
         & pTipo & "' and cod_preanalisis = '" & mExpediente & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
