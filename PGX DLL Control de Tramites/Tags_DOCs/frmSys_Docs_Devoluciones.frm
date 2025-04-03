VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmSys_Docs_Devoluciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Documentos: Devoluciones"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   11895
      _Version        =   1572864
      _ExtentX        =   20981
      _ExtentY        =   11668
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   11895
      _Version        =   1572864
      _ExtentX        =   20981
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin XtremeSuiteControls.ComboBox cboTipodoc 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   3975
      _Version        =   1572864
      _ExtentX        =   7011
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmSys_Docs_Devoluciones.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      ToolTipText     =   "Exportar"
      Top             =   1320
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Picture         =   "frmSys_Docs_Devoluciones.frx":0727
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Documento:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Devolución de Documentos"
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
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSys_Docs_Devoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim mTagDevolucion As String, mTagAplicado As String
Dim vTipoDoc As String


Private Sub btnAplicar_Click()
 Call sbAplicarRecepcionDevolucion
End Sub



Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Call Excel_Exportar_Lsw(lsw, PrgBar)

PrgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboTipodoc_Click()
If vPaso Then Exit Sub
If cboTipodoc.ListCount < 0 Then Exit Sub

Call sbCargaInformacion

End Sub

Private Sub Form_Activate()
vModulo = 8

End Sub

Private Sub Form_Load()

vModulo = 8

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Tipo", 1800
    .Add , , "Identificación", 1800, vbCenter
    .Add , , "Nombre", 4500
    .Add , , "Usuario", 2800, vbCenter
    .Add , , "Fecha", 1800
End With

  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '11'"
  Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagAplicado = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If


  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '12'"
  Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    
    If Not mTagDevolucion = Empty Then
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagDevolucion = Empty
            MsgBox "El código de tag definido el los parámetros para la Recepción/Devolución  no existe"
        End If
    End If

vPaso = True
   strSQL = "select rtrim(Tipo_Documento) as IdX, rtrim(Descripcion) as Itmx" _
          & " from SIF_Documentos" _
          & " where Tipo_documento in('NC','ND','FND','FNC','CA', 'CD.Liq', 'BEAC', 'CBJ', 'FSL', 'REA', 'RH', 'TCP', 'TRFA', 'TCP', 'THCJ', 'TRA', 'THAV')" _
          & " order by Descripcion"
   Call sbCbo_Llena_New(cboTipodoc, strSQL, False, True)
vPaso = False

Call cboTipodoc_Click


End Sub

Private Sub sbCargaInformacion()

On Error GoTo vError

    Me.MousePointer = vbHourglass
    
   strSQL = "Select Top 300 T.COD_TRANSACCION,T.TIPO_DOCUMENTO,T.CLIENTE_IDENTIFICACION,T.CLIENTE_NOMBRE,T.REGISTRO_USUARIO,T.REGISTRO_FECHA" _
           & " from SIF_TRANSACCIONES T " _
           & " where TIPO_DOCUMENTO = '" & cboTipodoc.ItemData(cboTipodoc.ListIndex) & "' and T.ANALISTA_RECEPCION = 1" _
           & " Order by T.REGISTRO_FECHA desc"
   lsw.ListItems.Clear
   
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
         Set itmX = lsw.ListItems.Add(, , rs!Cod_Transaccion)
        itmX.SubItems(1) = rs!Tipo_Documento
        itmX.SubItems(2) = rs!CLIENTE_IDENTIFICACION & ""
        itmX.SubItems(3) = rs!CLIENTE_NOMBRE & ""
        itmX.SubItems(4) = rs!registro_usuario
        itmX.SubItems(5) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
        rs.MoveNext
    Loop
    rs.Close

Me.MousePointer = vbDefault
Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAplicarRecepcionDevolucion()

Dim i As Long, pTipoDoc As String

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar esta etiqueta", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If mTagDevolucion = Empty Then
    MsgBox "No se puede realizar el proceso no está definido la etiqueta de devolución"
    Exit Sub
End If


Me.MousePointer = vbHourglass

pTipoDoc = cboTipodoc.ItemData(cboTipodoc.ListIndex)

PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lsw.ListItems

For i = 1 To .Count
        
    If .Item(i).Checked Then
        Call sbSIFRegistraTags(pTipoDoc, mTagDevolucion, "Devolución la documentación de la liquidación", .Item(i).Text, "DOC" _
                              , pTipoDoc, .Item(i).Text)
        
        strSQL = "update SIF_TRANSACCIONES set analista_recepcion = 0" _
               & "  where cod_transaccion = '" & .Item(i).Text & "'  and  tipo_documento = '" & pTipoDoc & "' "
        
        Call ConectionExecute(strSQL)
    End If
    
    PrgBar.Value = PrgBar.Value + 1

Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Proceso concluído con éxito!", vbInformation
Call sbCargaInformacion

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Long

    fxValidaNoDuplicados = False

'    For i = 1 To lsw.ListItems.Count
'
'        If lsw.ListItems(i).Text = Trim(txtCodigo.Text) Then
'            fxValidaNoDuplicados = True
'        End If
'
'    Next i

End Function

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub
