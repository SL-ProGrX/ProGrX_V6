VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_Seleccionar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7032
   ClientLeft      =   48
   ClientTop       =   0
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7032
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5172
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5892
      _Version        =   1245187
      _ExtentX        =   10393
      _ExtentY        =   9123
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   384
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   5892
      _Version        =   1245187
      _ExtentX        =   10393
      _ExtentY        =   677
      _StockProps     =   79
      Caption         =   "Cerrar el Sistema"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5640
      Top             =   480
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
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
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   4932
      _Version        =   1245187
      _ExtentX        =   8700
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
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   852
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5892
      _Version        =   1245187
      _ExtentX        =   10393
      _ExtentY        =   1503
      _StockProps     =   14
      Caption         =   "Seleccione la contabilidad que desea administrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
End
Attribute VB_Name = "frmCntX_Seleccionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxExistenContabilidades() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Contabilidades"
Call OpenRecordSet(rsX, strSQL, 0)
fxExistenContabilidades = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function



Private Sub btnCerrar_Click()
  End 'Cerrar el Sistema
End Sub

Private Sub Form_Load()

With lsw.ColumnHeaders
  .Clear
  .Add , , "Código", 1000
  .Add , , "Descripción", 4500
End With

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 Call sbSelecciona(Item.Text)
End Sub

Private Sub Timer1_Timer()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListViewItem
Dim Ix As Integer, lngEmpresa As Long, vFecha As Date

Set Me.Icon = frmContenedor.Icon

Timer1.Interval = 0

On Error Resume Next

  Ix = 1
  Do While Ix < 3 And Not fxExistenContabilidades
    MsgBox "No Existen Contabilidades Creadas, Necesita Crear al menos Una...", vbInformation
    frmCntX_Contabilidades.Show vbModal
    Ix = Ix + 1
  Loop

If Ix > 1 Then
  If Not fxExistenContabilidades Then
    MsgBox "No se creó ninguna contabilidad, no se permiten más intentos, reinicie el ProGrX: Contabilidad y vuelva a intentarlo...", vbCritical
    End
  End If
End If


'1. Primero Buscar los datos de la ultima utilización del usuario
'2. Si no tiene, entonces continuar
vFecha = fxFechaServidor

If gCntX_Parametros.MuestraTodas Then
   
   strSQL = "select cod_contabilidad,nombre from CntX_Contabilidades"
   Call OpenRecordSet(rs, strSQL, 0)
   
   gCntX_Parametros.PeriodoAnio = Year(vFecha)
   gCntX_Parametros.PeriodoMes = Month(vFecha)
     Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(lsw.ListItems.Count + 1, , rs!COD_CONTABILIDAD)
          itmX.SubItems(1) = rs!Nombre
      rs.MoveNext
     Loop
     txtCodigo.SetFocus
     
     If rs.RecordCount = 1 Then
        rs.MoveFirst
        lngEmpresa = rs!COD_CONTABILIDAD
        rs.Close
        strSQL = "select * from CntX_Acceso_Historico where usuario = '" & glogon.Usuario _
               & "' and cod_contabilidad = " & lngEmpresa
        Call OpenRecordSet(rs, strSQL, 0)
        If Not rs.EOF And Not rs.BOF Then
            gCntX_Parametros.PeriodoAnio = rs!Anio
            gCntX_Parametros.PeriodoMes = rs!Mes
            gCntX_Parametros.CodigoConta = rs!COD_CONTABILIDAD
            rs.Close
            Call sbSelecciona(gCntX_Parametros.CodigoConta)
        Else
            Call sbSelecciona(lngEmpresa)
        End If
     End If
Else
   
   strSQL = "select * from CntX_Acceso_Historico where usuario = '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL, 0)
   If Not rs.EOF And Not rs.BOF Then
       gCntX_Parametros.PeriodoAnio = rs!Anio
       gCntX_Parametros.PeriodoMes = rs!Mes
       gCntX_Parametros.CodigoConta = rs!COD_CONTABILIDAD
       rs.Close
       Call sbSelecciona(gCntX_Parametros.CodigoConta)
   Else
       gCntX_Parametros.PeriodoAnio = Year(vFecha)
       gCntX_Parametros.PeriodoMes = Month(vFecha)
   End If

End If

End Sub


Private Sub txtCodigo_Change()
If IsNumeric(txtCodigo) Then
  txtDescripcion = fxEmpresa(txtCodigo)
End If
End Sub


Private Function fxEmpresa(vCodigo As Long) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "Select Nombre from CntX_Contabilidades where cod_contabilidad = " & vCodigo
Call OpenRecordSet(rsX, strSQL)

If Not rsX.EOF And Not rsX.BOF Then fxEmpresa = rsX!Nombre
rsX.Close

End Function

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  If txtDescripcion <> "" And IsNumeric(txtCodigo) Then
    Call sbSelecciona(txtCodigo)
  End If
End If
End Sub


Private Sub sbSelecciona(vCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

'strSQL = "select * from CntX_Contabilidades where cod_contabilidad = " & vCodigo

strSQL = "select C.*" _
       & " from CntX_Contabilidades C inner join CNTX_CONTA_USUARIOS U on C.cod_contabilidad = U.cod_contabilidad" _
       & " and U.usuario = '" & glogon.Usuario & "'" _
       & " where C.cod_contabilidad  = " & vCodigo


Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "Esta Contabilidad a sido Eliminada o no tiene Acceso a ella, verifique...", vbCritical
  Exit Sub
End If
 
 gCntX_Parametros.CodigoConta = rs!COD_CONTABILIDAD
 gCntX_Parametros.NombreEmpresa = rs!Nombre
 gCntX_Parametros.Nivel1 = rs!Nivel1
 gCntX_Parametros.Nivel2 = rs!Nivel2
 gCntX_Parametros.Nivel3 = rs!Nivel3
 gCntX_Parametros.Nivel4 = rs!Nivel4
 gCntX_Parametros.Nivel5 = rs!Nivel5
 
 gCntX_Parametros.Nivel6 = rs!Nivel6
 gCntX_Parametros.Nivel7 = rs!Nivel7
 gCntX_Parametros.Nivel8 = rs!Nivel8
 
 gCntX_Parametros.TotalChr = rs!Nivel1 + rs!Nivel2 + rs!Nivel3 + rs!Nivel4 + rs!Nivel5 + rs!Nivel6 + rs!Nivel7 + rs!Nivel8
 gCntX_Parametros.Mascara = fxCntX_CuentaMascara(rs!Nivel1, rs!Nivel2, rs!Nivel3, rs!Nivel4, rs!Nivel5, rs!Nivel6, rs!Nivel7, rs!Nivel8)
 gCntX_Parametros.MascaraCod = rs!Nivel1 & rs!Nivel2 & rs!Nivel3 & rs!Nivel4 & rs!Nivel5 & rs!Nivel6 & rs!Nivel7 & rs!Nivel8
rs.Close
 
  

 'Saca el Periodo Pendiente
strSQL = "select dbo.fxCntX_PeriodoActual(" & vCodigo & ") as 'Periodo'"
Call OpenRecordSet(rs, strSQL, 0)

    gCntX_Parametros.PeriodoAnio = Year(rs!Periodo)
    gCntX_Parametros.PeriodoMes = Month(rs!Periodo)

rs.Close

 Me.Hide


'Dim frmX As MDIForm
'
'For Each frmX In Forms
'   If Mid(frmX.Name, 1, 3) = "MDI" Then
'        Exit For
'   End If
'Next
'
'If frmX Is Nothing Then
'   Set frmX = MDI
'End If
' frmX.Show

vError:
 Unload frmCntX_Seleccionar

End Sub

Private Sub txtDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select C.cod_contabilidad,C.nombre" _
       & " from CntX_Contabilidades C inner join CNTX_CONTA_USUARIOS U on C.cod_contabilidad = U.cod_contabilidad" _
       & " and U.usuario = '" & glogon.Usuario & "'" _
       & " where C.nombre like '%" & txtDescripcion.Text & "%'"
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_CONTABILIDAD)
     itmX.SubItems(1) = rs!Nombre
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault
vError:

End Sub
