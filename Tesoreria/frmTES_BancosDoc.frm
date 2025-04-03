VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_BancosDoc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Documentos x Bancos"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   8280
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1932
      Left            =   1800
      TabIndex        =   11
      Top             =   1320
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9123
      _ExtentY        =   3408
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
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkDocAuto 
      Height          =   252
      Left            =   3600
      TabIndex        =   8
      Top             =   4080
      Width           =   3372
      _Version        =   1310723
      _ExtentX        =   5948
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Consecutivo automático   "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBorrar 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   6840
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Borrar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_BancosDoc.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1800
      TabIndex        =   4
      Top             =   600
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9128
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   6840
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2990
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Guardar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_BancosDoc.frx":07CD
   End
   Begin XtremeSuiteControls.CheckBox chkModConsec 
      Height          =   252
      Left            =   3600
      TabIndex        =   9
      Top             =   4440
      Width           =   3372
      _Version        =   1310723
      _ExtentX        =   5948
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Modifica Consecutivo?    "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboComprobante 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   3600
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9128
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
   Begin XtremeSuiteControls.CheckBox chkAutorizacion 
      Height          =   252
      Left            =   3360
      TabIndex        =   12
      Top             =   5880
      Width           =   3612
      _Version        =   1310723
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Requiere de Proceso de Autorización"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkEmision 
      Height          =   252
      Left            =   3360
      TabIndex        =   13
      Top             =   6240
      Width           =   3612
      _Version        =   1310723
      _ExtentX        =   6371
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Requiere de Proceso de Emisión"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtConsec 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   5640
      TabIndex        =   15
      Top             =   4920
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConsecInterno 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5130
         SubFormatType   =   1
      EndProperty
      Height          =   312
      Left            =   5640
      TabIndex        =   16
      Top             =   5280
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   550
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   1800
      TabIndex        =   17
      Top             =   240
      Width           =   5172
      _Version        =   1310723
      _ExtentX        =   9128
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Banco ..:"
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
      Height          =   312
      Index           =   0
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Consecutivo Interno"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   3600
      TabIndex        =   14
      Top             =   5280
      Width           =   2172
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta ..:"
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
      Height          =   312
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Consecutivo actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   4920
      Width           =   2172
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Label lblTipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      Top             =   3276
      Width           =   5172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_BancosDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbLlenaLsw(vBanco As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX  As ListViewItem

Me.MousePointer = vbHourglass

lblTipo.Tag = ""
lblTipo.Caption = ""

chkAutorizacion.Value = vbUnchecked
chkDocAuto.Value = vbUnchecked
chkModConsec.Value = vbUnchecked
chkEmision.Value = vbChecked

txtConsec.Text = "0"
txtConsecInterno.Text = "0"

cboComprobante.Text = "01 - Cheque Formula Continua"

lsw.ListItems.Clear
strSQL = "select D.*,A.tipo as TipoX" _
       & " from tes_tipos_doc D left join tes_banco_docs A on D.tipo = A.tipo" _
       & " and A.id_banco = " & vBanco & " order by A.tipo desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Tipo)
     itmX.SubItems(1) = rs!Descripcion
     
     itmX.Checked = IIf(IsNull(rs!TipoX), vbUnchecked, vbChecked)
     
     If itmX.Checked Then
            itmX.ForeColor = vbBlue
     End If
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub cbo_Click()
If cbo.ListCount = 0 Then Exit Sub
If vPaso Then Exit Sub
Call sbLlenaLsw(cbo.ItemData(cbo.ListIndex))
End Sub


Private Sub cboBanco_Click()
If vPaso Then Exit Sub
If cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

vPaso = True
strSQL = "select id_banco as 'IdX',descripcion as 'ItmX' from Tes_Bancos" _
       & " where estado = 'A' and cod_grupo = '" & cboBanco.ItemData(cboBanco.ListIndex) & "'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)
vPaso = False

Call cbo_Click

End Sub

Private Sub cboComprobante_Click()

If Mid(cboComprobante.Text, 1, 2) = "04" Then
   txtConsecInterno.Locked = False
Else
   txtConsecInterno.Locked = True
End If

End Sub

Private Sub chkAutorizacion_Click()
If chkAutorizacion.Value = vbChecked Then
   chkEmision.Value = vbChecked
   chkEmision.Enabled = False
Else
   chkEmision.Enabled = True
End If
End Sub

Private Sub cmdBorrar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If lblTipo.Tag = "" Then Exit Sub

i = MsgBox("Esta seguro que desea borrar este Tipo de Documento", vbYesNo)

If i = vbYes Then
    'Verifica que no existan transacciones registradas.
    strSQL = "select count(*) as Existe from tes_Transacciones where tipo = '" & lblTipo.Tag _
           & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
    
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
        MsgBox "Existen (" & rs!Existe & ") Transacciones registradas a este tipo de documento. NO SE PUEDE ELIMINAR", vbExclamation
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    'Elimina la asignación de usuarios a este tipo de documento
    strSQL = "delete tes_documentos_asg where tipo = '" & lblTipo.Tag _
           & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
    
    'Elimina la asignación del documento al banco
    strSQL = strSQL & Space(10) & "delete tes_banco_docs where tipo = '" & lblTipo.Tag _
           & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
    Call ConectionExecute(strSQL)
    
    
    Call Bitacora("Elimina", "Cta. Id: " & cbo.ItemData(cbo.ListIndex) & ", Tipo Doc: & " & lblTipo.Tag)
    
    MsgBox "Tipo de documento: " & lblTipo.Caption & " eliminado satisfactoriamente!", vbInformation
    
    Call cbo_Click
End If 'Si

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If lblTipo.Tag = "" Then Exit Sub

strSQL = "select isnull(count(*),0) as Existe from tes_banco_docs where tipo = '" _
       & lblTipo.Tag & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   strSQL = "insert tes_banco_docs(tipo,id_banco,reg_autorizacion,reg_emision,doc_auto,consecutivo,comprobante" _
          & ",mod_consec,cuenta_min,cuenta_max,CONSECUTIVO_DET, REGISTRO_FECHA, REGISTRO_USUARIO) values('" _
          & lblTipo.Tag & "'," & cbo.ItemData(cbo.ListIndex) & "," & chkAutorizacion.Value & "," & chkEmision.Value _
          & "," & chkDocAuto.Value & "," & txtConsec.Text & ",'" & Mid(cboComprobante.Text, 1, 2) & "'," _
          & chkModConsec.Value & ",0,22," & txtConsecInterno.Text _
          & ", dbo.mygetdate(),'" & glogon.Usuario & "')"

Else
  strSQL = "update tes_banco_docs set reg_autorizacion = " & chkAutorizacion.Value _
         & ",reg_emision = " & chkEmision.Value & ",mod_consec = " & chkModConsec.Value _
         & ",Doc_auto = " & chkDocAuto.Value & ",comprobante = '" & Mid(cboComprobante.Text, 1, 2) _
         & "', consecutivo = " & txtConsec & ",cuenta_min = 0,cuenta_max = 22" _
         & ", CONSECUTIVO_DET=  " & txtConsecInterno.Text & ", actualiza_fecha= dbo.mygetdate()" _
         & ",Actualiza_Usuario = '" & glogon.Usuario & "'" _
         & " where tipo = '" & lblTipo.Tag & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
End If
Call ConectionExecute(strSQL)

rs.Close

MsgBox "Documentos actualizados y asignados al Banco satisfactoriamente..."

Call cbo_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Tipo", 900
    .Add , , "Descripción", 3500
End With
lsw.Checkboxes = False

cboComprobante.Clear
cboComprobante.AddItem "01 - Cheque Formula Continua"
cboComprobante.AddItem "02 - Cheque Formula Block"
cboComprobante.AddItem "03 - Boleta de Transacción"
cboComprobante.AddItem "04 - Transferencia Electrónica"



vPaso = True

strSQL = "select COD_GRUPO as 'Idx', DESCRIPCION as 'ItmX'" _
       & " From TES_BANCOS_GRUPOS Where ACTIVO = 1"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

Call cboBanco_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

If lsw.ListItems.Count = 0 Then Exit Sub

lblTipo.Tag = Item.Text
lblTipo.Caption = Item.SubItems(1)

chkAutorizacion.Value = vbUnchecked
chkDocAuto.Value = vbUnchecked
cboComprobante.Text = "01 - Cheque Formula Continua"

strSQL = "select * from tes_banco_docs where id_banco = " & cbo.ItemData(cbo.ListIndex) _
       & " and tipo = '" & Item.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   chkAutorizacion.Value = rs!reg_autorizacion
   chkEmision.Value = rs!reg_emision
   chkDocAuto.Value = rs!doc_auto
   chkModConsec.Value = rs!mod_consec
  Select Case Trim(rs!comprobante)
    Case "01"
        cboComprobante.Text = "01 - Cheque Formula Continua"
    Case "02"
        cboComprobante.Text = "02 - Cheque Formula Block"
    Case "03"
        cboComprobante.Text = "03 - Boleta de Transacción"
    Case "04"
        cboComprobante.Text = "04 - Transferencia Electrónica"
  End Select
  
  txtConsec.Text = rs!Consecutivo
  txtConsecInterno.Text = rs!CONSECUTIVO_DET
  
End If
rs.Close

Call chkAutorizacion_Click

End Sub
