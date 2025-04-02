VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_DivisasDC 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Divisas: Ajustes por Diferencial Cambiario"
   ClientHeight    =   7080
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   8328
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3972
      Left            =   0
      TabIndex        =   10
      Top             =   2880
      Width           =   8292
      _Version        =   1245187
      _ExtentX        =   14626
      _ExtentY        =   7006
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
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   4
      Top             =   6948
      Width           =   8328
      _ExtentX        =   14690
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.PushButton cmdProcesar 
      Height          =   612
      Left            =   6600
      TabIndex        =   5
      Top             =   1680
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Procesar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCntX_DivisasDC.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   4452
      _Version        =   1245187
      _ExtentX        =   7853
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
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtTCCompra 
      Height          =   312
      Left            =   4560
      TabIndex        =   8
      Top             =   1680
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtTCVenta 
      Height          =   312
      Left            =   4560
      TabIndex        =   9
      Top             =   2040
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption lblPeriodo 
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   8292
      _Version        =   1245187
      _ExtentX        =   14626
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Asientos pendientes de Autorización"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio > Venta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Cambio > Compra"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCntX_DivisasDC.frx":09C3
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_DivisasDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If Not vPaso Then Exit Sub
 
strSQL = "SELECT Top 50 ID_Cambio,TC_Compra,TC_Venta,Inicio,Corte" _
       & " FROM CntX_Divisas_Tipo_Cambio where cod_divisa = '" & cbo.ItemData(cbo.ListIndex) _
       & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and datepart(month,corte) = " & gCntX_Parametros.PeriodoMes _
       & " and datepart(year,corte) = " & gCntX_Parametros.PeriodoAnio _
       & " order by corte desc"
Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_cambio)
     itmX.SubItems(1) = Format(rs!tc_compra, "Standard")
     itmX.SubItems(2) = Format(rs!tc_venta, "Standard")
     itmX.SubItems(3) = Format(rs!inicio, "dd/mm/yyyy")
     itmX.SubItems(4) = Format(rs!Corte, "dd/mm/yyyy")
 rs.MoveNext
Loop
rs.Close


End Sub

Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & strCuenta & "' and acepta_movimientos =1"

Call OpenRecordSet(rsX, strSQL, 0)
 fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function

Private Sub cmdProcesar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDivisa As String, curGanancia As Currency, curPerdida As Currency
Dim vCuentaIngresos As String, vCuentaGastos As String, vAsientoNum As String
Dim curTC As Currency, vAnio As Long, vMes As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
prgBar.Visible = True

vDivisa = cbo.ItemData(cbo.ListIndex)



vAnio = gCntX_Parametros.PeriodoAnio
vMes = gCntX_Parametros.PeriodoMes

'Cuentas a Utilizar
strSQL = "select * from Cntx_Divisas where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_Divisa = '" & vDivisa & "'"
Call OpenRecordSet(rs, strSQL, 0)
 vCuentaIngresos = Trim(rs!cod_cuenta & "")
 vCuentaGastos = Trim(rs!cod_cuenta_gasto & "")
rs.Close

'Validar Cuentas
If Not fxVerificaCuenta(vCuentaIngresos) Then
   MsgBox "La cuenta de Ingreso por Diferencial Cambiario no es válida, revise la configuración de la divisa!", vbExclamation
   Exit Sub
End If
If Not fxVerificaCuenta(vCuentaGastos) Then
   MsgBox "La cuenta de Gasto por Diferencial Cambiario no es válida, revise la configuración de la divisa!", vbExclamation
   Exit Sub
End If

If Not IsNumeric(txtTCCompra.Text) Or Not IsNumeric(txtTCVenta.Text) Then
   MsgBox "Especifique un Tipo de Cambio válido para la aplicación del Diferencial!", vbExclamation
   Exit Sub
End If

'Ejecuta el Proceso de Diferencial
strSQL = "exec spCntX_DiferencialCambiario " & gCntX_Parametros.CodigoConta & "," & vAnio & "," & vMes _
       & ",'" & vDivisa & "'," & CCur(txtTCCompra.Text) & "," & CCur(txtTCVenta.Text) & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL, 0)

Call Bitacora("Aplica", "Asientos-Diferencial Cambiario (Conta.:" & gCntX_Parametros.CodigoConta & " - Periodo.: " & vAnio & "-" & vMes _
            & " - Divisa.: " & vDivisa & ")")

Me.MousePointer = vbDefault
prgBar.Visible = False

Call sbLimpiaPantalla

MsgBox "Asientos por Diferencial Cambiario Aplicado Satisfactoriamente...", vbInformation

Exit Sub

vError:
  prgBar.Visible = False
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

 With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1200, vbCenter
    .Add , , "TC.Compra", 1800, vbRightJustify
    .Add , , "TC.Venta", 1800, vbRightJustify
    .Add , , "Inicio", 1800, vbCenter
    .Add , , "Corte", 1800, vbCenter
 End With
 


vPaso = False

strSQL = "select rtrim(cod_divisa) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " From CntX_Divisas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and divisa_local = 0"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

vPaso = True
Call cbo_Click


 Call sbLimpiaPantalla
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()
Dim vPeriodoDesc As String

vPeriodoDesc = fxCntX_PeriodoDesc(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes)

txtTCCompra = ""
txtTCVenta = ""

lblPeriodo.Caption = vPeriodoDesc
End Sub

Private Sub lsw_Click()

If lsw.ListItems.Count = 0 Then Exit Sub

txtTCCompra = CCur(lsw.SelectedItem.SubItems(1))
txtTCVenta = CCur(lsw.SelectedItem.SubItems(2))

End Sub

Private Sub txtTCCompra_GotFocus()
On Error GoTo vError
 txtTCCompra = CCur(txtTCCompra)
vError:
End Sub

Private Sub txtTCCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdProcesar.SetFocus
End Sub

Private Sub txtTCCompra_LostFocus()
On Error GoTo vError
 txtTCCompra = Format(CCur(txtTCCompra), "Standard")
vError:
End Sub

Private Sub txtTCVenta_GotFocus()
On Error GoTo vError
 txtTCVenta = CCur(txtTCVenta)
vError:
End Sub

Private Sub txtTCVenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTCCompra.SetFocus
End Sub

Private Sub txtTCVenta_LostFocus()
On Error GoTo vError
 txtTCVenta = Format(CCur(txtTCVenta), "Standard")
vError:
End Sub
