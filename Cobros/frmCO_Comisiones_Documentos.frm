VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_Comisiones_Documentos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Comisiones: Tipos de Documentos Admitidos "
   ClientHeight    =   8148
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   9084
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8148
   ScaleWidth      =   9084
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6732
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9012
      _Version        =   1245187
      _ExtentX        =   15896
      _ExtentY        =   11874
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
      ShowBorder      =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Documentos admitidos en el cálculo"
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
      Height          =   480
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_Comisiones_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activate()
vModulo = 4

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
 .Clear
 .Add , , "Tipo", 1500
 .Add , , "Descripción", 3000
 .Add , , "Fecha", 1800
 .Add , , "Usuario", 1800
End With

Call sbList_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbList_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear


vPaso = True

strSQL = "select Ant.TIPO_DOCUMENTO,Ant.DESCRIPCION,isnull(Asg.TIPO_DOCUMENTO,'No-ASG') as 'Asignado'" _
       & ", Asg.Registro_Fecha, Asg.Registro_Usuario " _
       & " from SIF_DOCUMENTOS Ant left join CBR_COMISIONES_TDOC Asg on Ant.TIPO_DOCUMENTO = Asg.TIPO_DOCUMENTO" _
       & " ORDER BY isnull(Asg.TIPO_DOCUMENTO,'000') desc ,  Ant.TIPO_DOCUMENTO"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!TIPO_DOCUMENTO)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!Registro_Fecha & ""
      itmX.SubItems(3) = rs!Registro_Usuario & ""
      
      itmX.Tag = rs!TIPO_DOCUMENTO
      If Trim(rs!Asignado) <> "No-ASG" Then
         itmX.Checked = True
      End If
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



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String


If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CBR_COMISIONES_TDOC(TIPO_DOCUMENTO,registro_fecha,registro_usuario)" _
          & " values('" & Item.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete CBR_COMISIONES_TDOC where TIPO_DOCUMENTO = '" & Item.Tag & "'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

