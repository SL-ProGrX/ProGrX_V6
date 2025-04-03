VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmSIF_RecepcionDevolucionesNcNd 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Documentos: Devoluciones"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   1800
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
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   11895
      _Version        =   1572864
      _ExtentX        =   20981
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin XtremeSuiteControls.PushButton cmdAgregar 
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   1200
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
      Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   345
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipodoc 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
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
      Left            =   10680
      TabIndex        =   5
      Top             =   1200
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
      Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":0720
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1200
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
      TabIndex        =   2
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
Attribute VB_Name = "frmSIF_RecepcionDevolucionesNcNd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim mTagDevolucion As String
Dim mTagAplicado  As String
Dim vTipoDoc As String


Private Sub btnAplicar_Click()
 Call sbAplicarRecepcionDevolucion
End Sub

Private Sub cmdAgregar_Click()
If Trim(txtCodigo.Text) <> "" Then Call sbCargaInformacion
End Sub

Private Sub Form_Activate()
vModulo = 10

End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vModulo = 10

  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '11'"
  Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagAplicado = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If
    rs.Close


  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '12'"
  Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagDevolucion = Empty
            MsgBox "El código de tag definido el los parámetros para la Recepción/Devolución  no existe"
        End If
        rs.Close
        
    End If

   strSQL = "select rtrim(Tipo_Documento) + ' - ' + Descripcion as Itmx" _
          & " from SIF_Documentos" _
          & " where Tipo_documento in('NC','ND','FND','FNC','CA', 'CD.Liq', 'BEAC', 'CBJ', 'FSL', 'REA', 'RH', 'TCP', 'TRFA', 'TCP', 'THCJ', 'TRA', 'THAV')" _
          & " order by Tipo_Documento"
          
   Call sbLlenaCbo(cboTipodoc, strSQL, False, False)

End Sub

Private Sub optNc_Click()
If optNc = True Then vTipoDoc = "NC"
End Sub

Private Sub optNd_Click()
If optNd = True Then vTipoDoc = "ND"

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
     If optNc = True Then
        vTipoDoc = "NC"
     ElseIf optNd = True Then
        vTipoDoc = "ND"
     End If
     
        Call sbCargaInformacion
    End If
End Sub

Private Sub sbCargaInformacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem
On Error GoTo vError

    If Not IsNumeric(txtCodigo) Then
        Exit Sub
    End If
    
    If fxValidaNoDuplicados = True Then
        MsgBox "El documento se ya fue digitada"
        txtCodigo.Text = Empty
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    'Valida no agregar en forma mismo tag en forma consecutiva
'    strsql = "SELECT dbo.fxSIFValidaTagRev('" & Trim(txtCodigo) & "','" & Trim(mTagDevolucion) & "','" & Trim(mTagDevolucion) & "','04','" & vTipoDoc & "',NULL)"
'    rs.Open strsql, glogon.Conection, adOpenStatic
'    If Not rs.EOF Then
'        If rs.Fields(0) = 1 Then
'           MsgBox "No es posible registrar en forma consecutiva dos recepciones en el mismo docuemnto " & txtCodigo.Text
'           txtCodigo.Text = Empty
'           rs.Close
'           Exit Sub
'        End If
'    End If
'    rs.Close
    
   strSQL = "Select T.COD_TRANSACCION,T.TIPO_DOCUMENTO,T.CLIENTE_IDENTIFICACION,S.NOMBRE,T.REGISTRO_USUARIO,T.REGISTRO_FECHA" _
           & " from SIF_TRANSACCIONES T inner join Socios S on T.CLIENTE_IDENTIFICACION = S.cedula" _
           & " where TIPO_DOCUMENTO = '" & vTipoDoc & "' and T.ANALISTA_REVISION =1" _
           & " and T.COD_TRANSACCION in (select codigo from SIF_CONTROL_TAGS where codigo = '" & Trim(txtCodigo.Text) _
           & "' and TAG_CODIGO = 'S04' and cod_modulo ='06' ) and dbo.fxSIFValidaTagRev(T.COD_TRANSACCION,'" & Trim(mTagAplicado) _
           & "','" & Trim(mTagDevolucion) & "','06',T.TIPO_DOCUMENTO,NULL) <> 1"
 
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    If Not rs.EOF Then
         Set itmX = lswDocumento.ListItems.Add(, , rs!Cod_Transaccion)
        itmX.SubItems(1) = rs!Tipo_Documento
        itmX.SubItems(2) = rs!CLIENTE_IDENTIFICACION
        itmX.SubItems(3) = rs!Nombre
        itmX.SubItems(4) = rs!registro_usuario
        itmX.SubItems(5) = Format(rs!Registro_Fecha, "dd/mm/yyyyy")
    End If
    rs.Close

    txtCodigo.Text = Empty
    txtCodigo.SetFocus

    Exit Sub
    
vError:
        MsgBox Err.Description

End Sub

Private Sub sbAplicarRecepcionDevolucion()
Dim i As Integer, strSQL As String

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar esta etiqueta", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If mTagDevolucion = Empty Then
    MsgBox "No se puede realizar el proceso no está definido la etiqueta de devolución"
    Exit Sub
End If


Me.MousePointer = vbHourglass

PrgBar.Max = lswDocumento.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswDocumento.ListItems

For i = 1 To .Count
        Call sbSIFRegistraTags(.Item(i).Text, mTagDevolucion, "Recepción de Devolución la documentación de la afiliación", .Item(i).SubItems(3), "04")
        strSQL = "update SIF_TRANSACCIONES set analista_recepcion  =0 where cod_transaccion = '" & .Item(i).Text & "'  and  tipo_documento = '" & vTipoDoc & "' "
        glogon.Conection.Execute strSQL
    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso concluido con éxito...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lswDocumento.ListItems.Count

        If lswDocumento.ListItems(i).Text = Trim(txtCodigo.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function






