VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGenEnlacesCredito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Códigos de Enlace con Sistema de Crédito"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7704
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   7704
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodCredito 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   940
   End
   Begin VB.TextBox txtDescripcion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   5480
   End
   Begin VB.TextBox txtCodInst 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   940
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
      _ExtentX        =   13145
      _ExtentY        =   7218
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod.Ins"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   9763
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Código"
         Object.Width           =   1658
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmGenEnlacesCredito.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmGenEnlacesCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbCargaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

strSQL = "select I.cod_institucion,I.descripcion,P.cod_credito" _
       & " from instituciones I inner join PV_PARINSTITUCIONES P" _
       & " on I.cod_institucion = P.cod_institucion"
Call OpenRecordSet(rs, strSQL, 0)
lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_institucion)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = rs!cod_credito
 rs.MoveNext
Loop
rs.Close

vError:



End Sub

Private Sub Form_Activate()
vModulo = 34
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 34
Call Formularios(Me)

'Actualiza la Tabla de Enlaces de Credito con Instituciones no configuradas
'o Nuevas
strSQL = "INSERT INTO PV_PARINSTITUCIONES(COD_INSTITUCION,COD_CREDITO)" _
       & " (SELECT COD_INSTITUCION,'' FROM INSTITUCIONES" _
       & " WHERE COD_INSTITUCION NOT IN(SELECT COD_INSTITUCION FROM PV_PARINSTITUCIONES))"
Call ConectionExecute(strSQL)

Call sbCargaLsw

vError:

Call RefrescaTags(Me)

End Sub

Private Sub lsw_Click()

If lsw.ListItems.Count > 0 Then
  txtCodInst = lsw.SelectedItem
  txtDescripcion = lsw.SelectedItem.SubItems(1)
  txtCodCredito = lsw.SelectedItem.SubItems(2)
  txtCodCredito.SetFocus
End If

End Sub

Private Sub txtCodCredito_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

If KeyCode = vbKeyF4 And txtCodInst <> "" Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "Select Codigo,Descripcion from catalogo"
   gBusquedas.Filtro = " And cod_institucion = " & txtCodInst
   gBusquedas.Convertir = "N"
   frmBusquedas.Show vbModal
   txtCodCredito = gBusquedas.Resultado
End If

If KeyCode = vbKeyReturn And txtCodInst <> "" Then
   strSQL = "update PV_PARINSTITUCIONES set cod_credito = '" & Trim(txtCodCredito) _
          & "' where cod_institucion = " & txtCodInst
   Call ConectionExecute(strSQL)
   Call sbCargaLsw
   txtCodCredito = ""
   txtCodInst = ""
   txtDescripcion = ""
End If

End Sub
