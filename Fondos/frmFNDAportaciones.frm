VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFNDAportaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aportes a Contratos"
   ClientHeight    =   4950
   ClientLeft      =   1125
   ClientTop       =   2565
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAporte 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox cboTipoDoc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmFNDAportaciones.frx":0000
      Left            =   4200
      List            =   "frmFNDAportaciones.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      Picture         =   "frmFNDAportaciones.frx":0027
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   4695
   End
   Begin VB.TextBox txtOperadora 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtContrato 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1815
      Left            =   1320
      TabIndex        =   19
      Top             =   1800
      Width           =   7695
      _Version        =   524288
      _ExtentX        =   13573
      _ExtentY        =   3201
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDAportaciones.frx":015F
      VScrollSpecialType=   2
      AppearanceStyle =   0
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2880
      TabIndex        =   18
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2880
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   9000
      X2              =   0
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8880
      X2              =   0
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label6 
      Caption         =   "SubCuentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   6360
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
End
Attribute VB_Name = "frmFNDAportaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
vError:
End Sub


Private Sub CmdAplicar_Click()
Dim vRecibo As Long, vFecha As Date, i As Integer
Dim vProceso As Long, vCombo As String, curMonto As Currency

Dim x As New clsImpresoras
Dim vDriver, vTipo As String

Dim strSQL As String, vConcepto As String
Dim vTipoDoc As String, vTcon As String

On Error GoTo vError

If txtAporte.Tag = "N" Then
  MsgBox "Este plan no permite movimientos en Cajas, verifique...", vbExclamation
  Exit Sub
End If

If txtEstado.Tag = "L" Then
  MsgBox "Este contrato se encuentra Liquidado, verifique...", vbExclamation
  Exit Sub
End If

If CCur(txtAporte) = 0 Then
  MsgBox "No se especificó ningún aporte, verifique...", vbExclamation
  Exit Sub
End If

If Trim(txtContrato) = "" Or Trim(txtAporte) = "" Or Trim(cboTipoDoc) = "" Then
 MsgBox "Faltan Datos", vbExclamation, "No se puede aplicar"
Else

 vCombo = IIf(Trim(cboTipoDoc) = "Orden Inversion", "Nota Debito", Trim(cboTipoDoc))
 vConcepto = "FND001"
 vFecha = fxFechaServidor
 vProceso = Year(vFecha) & Format(Month(vFecha), "00")
 
 If GLOBALES.SysDocVersion = 1 Then
   vTipoDoc = fxgFNDTipoASEDoc(vCombo)
   vTcon = fxgFNDTipoDoc(vCombo)
   vRecibo = fxgFNDDocumento(Trim(vCombo), gFondos.Operadora, Trim(gFondos.Plan), Trim(txtContrato), CCur(txtAporte), "")
 Else
   vTipoDoc = fxgFNDTipoASEDoc(vCombo)
   vTcon = vTipoDoc
   vRecibo = fxgFNDDocumento(vTipoDoc, gFondos.Operadora, Trim(gFondos.Plan), Trim(txtContrato), CCur(txtAporte), vConcepto)
 End If
 
 If vRecibo = 0 Then
    MsgBox "Documentación Especificada no es válida ...", vbExclamation
    Exit Sub
 End If
 
 glogon.Conection.BeginTrans
 
 strSQL = "Insert into fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
        & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,cod_concepto,usuario) Values(" & gFondos.Operadora & ",'" _
        & Trim(gFondos.Plan) & "'," & gFondos.Contrato & ",dbo.MyGetdate()," _
        & CCur(txtAporte) & "," & vProceso & ",'" & vTcon & "','" & vRecibo & "','" & vConcepto & "','" & glogon.Usuario & "')"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Registra", vTipoDoc & " Ope:" & gFondos.Operadora & " Plan:" & Trim(gFondos.Plan) & " Cont:" & Trim(txtContrato) & " Monto:" & Trim(txtAporte))
 
 strSQL = "Update Fnd_contratos set Aportes = Aportes + " & CCur(txtAporte) _
        & " where cod_operadora=" & gFondos.Operadora _
        & " and cod_plan='" & Trim(gFondos.Plan) & "'" _
        & " and cod_contrato=" & gFondos.Contrato
 Call ConectionExecute(strSQL)
 
 
 If txtAporte.Locked Then
  For i = 1 To vGrid.MaxRows
     vGrid.col = 4
     vGrid.Row = i
     
     curMonto = CCur(vGrid.Text)
     vGrid.col = 1
     If curMonto > 0 Then
     
        strSQL = "Insert into fnd_SubCuentas_detalle(idx,Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon) Values(" & vGrid.Text & "," & gFondos.Operadora & ",'" _
               & Trim(gFondos.Plan) & "'," & gFondos.Contrato & ",dbo.MyGetdate()," _
               & curMonto & "," & vProceso & ",'" & vTcon & "','" & vRecibo & "')"
        Call ConectionExecute(strSQL)
        
        strSQL = "Update Fnd_subCuentas set Aportes = Aportes + " & curMonto _
               & " where cod_operadora=" & gFondos.Operadora _
               & " and cod_plan='" & Trim(gFondos.Plan) & "'" _
               & " and cod_contrato=" & gFondos.Contrato & " and IdX = " & vGrid.Text
        Call ConectionExecute(strSQL)
     
     End If
  Next i
 End If
 
 
 glogon.Conection.CommitTrans
 
 If GLOBALES.SysDocVersion = 1 Then
     Call sbgFNDImprimeRecibo(vRecibo, vTipoDoc, gFondos.Operadora)
 Else
     Call sbImprimeRecibo(vRecibo, vTipoDoc)
 End If
 
 MsgBox "Aporte aplicado, con : " & vCombo & " # " & vRecibo, vbInformation
 Call sbLimpiaPantalla

End If

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  glogon.Conection.RollbackTrans
End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
'Me.Icon = frmFNDOperadoras.Icon
 
vModulo = 18 'Fondo de Inversion
vGrid.AppearanceStyle = fxGridStyle

Call sbLimpiaPantalla
'Call Formularios(Me)
'Call RefrescaTags(Me)

cboTipoDoc.Text = "Recibo"
End Sub


Private Sub sbLimpiaPantalla()
Dim strSQL As String, rs As New ADODB.Recordset

txtAporte = 0
txtAporte.Tag = "S"
txtAporte.Locked = False

vGrid.MaxRows = 0

strSQL = "select C.cedula,S.nombre,P.descripcion as PlanX,O.descripcion as OperadoraX" _
       & ",C.cod_plan,C.cod_contrato,C.cod_operadora,C.estado,C.fecha_Inicio,isnull(P.cuenta_Maestra,0) as CuentaMaestra" _
       & ",P.Tipo_CDP,C.Inversion,P.Permite_Mov_Cajas,C.aportes" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and C.cod_operadora = P.cod_operadora" _
       & " inner join fnd_operadoras O on C.cod_operadora = O.cod_operadora" _
       & " where C.cod_operadora = " & gFondos.Operadora & " and C.cod_plan = '" & gFondos.Plan _
       & "' and C.cod_Contrato = " & gFondos.Contrato
Call OpenRecordSet(rs, strSQL)
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 txtOperadora = rs!operadoraX
 txtDescripcion = rs!PlanX
 txtContrato = rs!COD_CONTRATO
 txtFecha = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
 txtEstado.Tag = rs!Estado
 txtEstado = IIf((rs!Estado = "A"), "Activo", "Liquidado")

 If rs!tipo_cdp = 1 Then
    txtAporte.Locked = True
    If rs!aportes = 0 Then
       txtAporte = Format(rs!inversion, "Standard")
    End If
 End If
 
 
 If rs!CuentaMaestra = 1 Then
    txtAporte.Locked = True
    strSQL = "select IDx,Cedula,Nombre,0 from fnd_subCuentas where cod_operadora = " & gFondos.Operadora _
           & " and cod_plan = '" & gFondos.Plan & "' and cod_contrato = " & gFondos.Contrato _
           & " and estado = 'A'"
   Call sbCargaGrid(vGrid, 4, strSQL)
   vGrid.MaxRows = vGrid.MaxRows - 1
 End If
 
 If rs!PERMITE_MOV_CAJAS = 0 Then
  txtAporte.Tag = "N"
 End If
rs.Close

End Sub



Private Sub txtAporte_GotFocus()
On Error GoTo vError
 txtAporte = CCur(txtAporte)
vError:
End Sub

Private Sub txtAporte_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8, 46
  Case vbKeyReturn
     cboTipoDoc.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub


Private Sub txtAporte_LostFocus()
On Error GoTo vError
 txtAporte = Format(CCur(txtAporte), "Standard")
vError:
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAporte.SetFocus
End Sub


Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim i As Integer, curMonto As Currency

If col = 4 Then
 curMonto = 0
 For i = 1 To vGrid.MaxRows
   vGrid.col = 4
   vGrid.Row = i
   curMonto = curMonto + CCur(vGrid.Text)
 Next i
 
 txtAporte = Format(curMonto, "Standard")
 
End If

End Sub
