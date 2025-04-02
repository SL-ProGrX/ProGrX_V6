VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmSYS_APL_Dominios 
   Caption         =   "APL: Dominios Vinculados"
   ClientHeight    =   6072
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6072
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3732
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   9564
      _Version        =   1245185
      _ExtentX        =   16870
      _ExtentY        =   6583
      _StockProps     =   77
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   5880
      Top             =   1560
   End
   Begin XtremeSuiteControls.ComboBox cboDominio 
      Height          =   312
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5172
      _Version        =   1245185
      _ExtentX        =   9123
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administración de Dominios vinculados"
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
      Height          =   612
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   7932
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dominios vinculados:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dominio (Base):"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2172
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmSYS_APL_Dominios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboDominio_Click()

If vPaso Then Exit Sub
If cboDominio.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

strSQL = "exec spAPL_Dominios_Vinculados_List '" & cboDominio.ItemData(cboDominio.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

vPaso = True

Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Cod_Dominio)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!registro_fecha & ""
      itmX.SubItems(3) = rs!registro_usuario & ""
      itmX.Checked = IIf((rs!Asignado = 1), True, False)
  
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 38
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lsw.ColumnHeaders
  .Add , , "Dominio", 2000, vbCenter
  .Add , , "Descripción", 4000
  .Add , , "Fecha", 2000, vbCenter
  .Add , , "Usuario", 2000, vbCenter
End With

vPaso = True

strSQL = "exec spAPL_Dominios_Vinculados '" & gAPL.APL_Dominio & "'"
Call sbCbo_Llena_New(cboDominio, strSQL, False, True)

vPaso = False


'Call Formularios(Me)
'Call RefrescaTags(Me)


End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert APL_DOMINICIOS_VINCULADOS(COD_DOMINIO,COD_DOMINIO_VINCULADO, REGISTRO_FECHA, REGISTRO_USUARIO)" _
           & " VALUES('" & cboDominio.ItemData(cboDominio.ListIndex) & "','" & Item.Text & "',dbo.myGetdate(),'" & glogon.Usuario & "') "
    Item.SubItems(2) = Date
    Item.SubItems(3) = glogon.Usuario
    
Else
    strSQL = "delete APL_DOMINICIOS_VINCULADOS" _
           & " where COD_DOMINIO = '" & cboDominio.ItemData(cboDominio.ListIndex) & "' AND COD_DOMINIO_VINCULADO = '" & Item.Text & "'"
    Item.SubItems(2) = ""
    Item.SubItems(3) = ""
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call cboDominio_Click

End Sub
