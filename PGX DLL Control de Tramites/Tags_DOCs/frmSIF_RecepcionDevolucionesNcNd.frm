VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSIF_RecepcionDevolucionesNcNd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recepcion Devoluciones ND/NC"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":0000
   ScaleHeight     =   8085
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optNc 
      Caption         =   "Nota Crédito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.OptionButton optNd 
      Caption         =   "Nota Débito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   840
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   495
   End
   Begin MSComctlLib.ListView lswDocumento 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10186
      View            =   3
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cédula"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Usuario Reg."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fecha reg."
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAplicar 
      Height          =   570
      Left            =   10680
      TabIndex        =   3
      Top             =   7440
      Width           =   1305
      _ExtentX        =   2302
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
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9960
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionDevolucionesNcNd.frx":13916
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo Documento ..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label LblOperacion 
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Recepción de Devoluciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmSIF_RecepcionDevolucionesNcNd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTagDevolucion As String
Dim mTagAplicado  As String
Dim vTipoDoc As String


Private Sub cmdAgregar_Click()
If Trim(txtCodigo.Text) <> "" Then Call sbCargaInformacion
End Sub

Private Sub Form_Activate()
vModulo = 10

End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim rs As New ADODB.Recordset

vModulo = 10

  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '11'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF Then
        mTagAplicado = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If
    rs.Close



  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '12'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        rs.Open strSQL, glogon.Conection, adOpenStatic
        If rs.Fields(0) = 0 Then
            mTagDevolucion = Empty
            MsgBox "El código de tag definido el los parámetros para la Recepción/Devolución  no existe"
        End If
        rs.Close
        
    End If

End Sub

Private Sub optNc_Click()
If optNc = True Then vTipoDoc = "NC"
End Sub

Private Sub optNd_Click()
If optNd = True Then vTipoDoc = "ND"

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Call sbAplicarRecepcionDevolucion
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






