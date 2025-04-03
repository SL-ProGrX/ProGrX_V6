VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmFNDAnulaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulaciones a Contratos"
   ClientHeight    =   5055
   ClientLeft      =   1125
   ClientTop       =   2640
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbAnulacion 
      Height          =   1212
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Width           =   9132
      _Version        =   1441793
      _ExtentX        =   16108
      _ExtentY        =   2138
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdAnular 
         Height          =   732
         Left            =   7440
         TabIndex        =   16
         Top             =   240
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Anular"
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
         Appearance      =   17
         Picture         =   "frmFNDAnulaciones.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtAporte 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   5160
         TabIndex        =   18
         Top             =   480
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Anulación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   3600
         TabIndex        =   17
         Top             =   480
         Width           =   1332
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1812
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   8052
      _Version        =   524288
      _ExtentX        =   14203
      _ExtentY        =   3196
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmFNDAnulaciones.frx":0995
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   480
      Width           =   4692
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOperadora 
      Height          =   312
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   4692
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   312
      Left            =   1680
      TabIndex        =   10
      Top             =   1080
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1680
      TabIndex        =   11
      Top             =   1440
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   312
      Left            =   7680
      TabIndex        =   13
      Top             =   120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   312
      Left            =   7680
      TabIndex        =   14
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
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
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
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
      Height          =   315
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Left            =   6720
      TabIndex        =   2
      Top             =   480
      Width           =   852
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SubCuentas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmFNDAnulaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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


Private Sub cmdAnular_Click()
Dim rs As New ADODB.Recordset, curSobrante As Currency
Dim vRecibo As Long, vFecha As Date, vProceso As Long
Dim vAporteActual As Currency, vRendiActual As Currency
Dim vAplAporte As Currency, vAplRendi As Currency
Dim strSQL As String, curMonto As Currency, i As Integer
Dim vTcon As String, vTipoDoc As String, vConcepto As String

On Error GoTo vError

If Trim(gFondos.Contrato) = "" Or Trim(txtAporte) = "" Or Trim(txtAporte) = "0" Or Not IsNumeric(txtAporte) Then
 MsgBox "No se especifico el contrato o el monto", vbExclamation, "No se puede aplicar"
 Exit Sub
End If
  
If txtAporte.Tag = "N" Then
 MsgBox "Este Plan no permite realizar movimientos en cajas, o se trata de un CDP", vbExclamation, "No se puede aplicar"
 Exit Sub
End If
  
  
strSQL = "select aportes,rendimiento from fnd_contratos where cod_operadora = " & gFondos.Operadora _
       & " and cod_plan = '" & gFondos.Plan & "' and cod_contrato = " & gFondos.Contrato
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  MsgBox "No se encontró el contrato...", vbExclamation
  Exit Sub
Else
  vAporteActual = IIf(IsNull(rs!APORTES), 0, rs!APORTES)
  vRendiActual = IIf(IsNull(rs!Rendimiento), 0, rs!Rendimiento)
End If
rs.Close

  
If CCur(txtAporte) > vAporteActual + vRendiActual Then
  MsgBox "La Anulación es mayor que el total de los aportes y rendimientos del contrato...", vbExclamation
  Exit Sub
End If
    
    
    
If fxFndParametro("01.2") = "S" Then
   strSQL = "exec spFndSeguridad_ApAnul " & gFondos.Operadora & ",'" & gFondos.Plan & "','" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
   If rs!Autoriza = 0 Then
        MsgBox "El Usuario no tiene nivel de Autorización para realizar este movimiento!", vbExclamation
        Exit Sub
   End If
End If
    
    
    
'Validar sub Contratos
If txtAporte.Locked Then
 For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 4
    If CCur(vGrid.Text) > 0 Then
       vGrid.col = 5
       curMonto = CCur(vGrid.Text)
       vGrid.col = 6
       curMonto = curMonto + CCur(vGrid.Text)
       vGrid.col = 4
       If CCur(vGrid.Text) > curMonto Then
            vGrid.col = 1
            MsgBox "La Anulación es mayor al total de los aportes y rendimientos de las subCuentas (" & vGrid.Text & ")...", vbExclamation
            Exit Sub
       End If
       
    End If
 Next i
End If
  
Me.MousePointer = vbHourglass


vFecha = fxFechaServidor
vProceso = Year(vFecha) & Format(Month(vFecha), "00")

'Distribuye Anulacion

curSobrante = CCur(txtAporte)

If curSobrante >= vRendiActual Then
   vAplRendi = vRendiActual
   curSobrante = curSobrante - vRendiActual
Else
   vAplRendi = curSobrante
   curSobrante = 0
End If


If curSobrante >= vAporteActual Then
   vAplAporte = vAporteActual
   curSobrante = curSobrante - vAporteActual
Else
   vAplAporte = curSobrante
   curSobrante = 0
End If
 

vConcepto = "FND002"
vRecibo = fxgFNDDocumento("FNC", gFondos.Operadora, Trim(gFondos.Plan), Trim(gFondos.Contrato), vAplAporte, vConcepto, vAplRendi)
vTipoDoc = "FNC"
vTcon = "FNC"


If vRecibo = 0 Then
   MsgBox "Datos del Documento no son válidos...", vbExclamation
   Exit Sub
End If
 
glogon.Conection.BeginTrans
 
strSQL = "Update Fnd_contratos set Aportes = Aportes - " & vAplAporte _
       & ",rendimiento = rendimiento - " & vAplRendi _
       & " where cod_operadora=" & gFondos.Operadora & " and cod_plan = '" & gFondos.Plan _
       & "' and cod_contrato = " & gFondos.Contrato
Call ConectionExecute(strSQL)
  
 
strSQL = "Insert into fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
       & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,cod_concepto,usuario,cod_Caja) Values(" & gFondos.Operadora & ",'" _
       & gFondos.Plan & "'," & gFondos.Contrato & ",dbo.MyGetdate()," & CCur(txtAporte) * -1 & "," _
       & vProceso & ",'" & vTcon & "','" & vRecibo & "','" & vConcepto & "','" & glogon.Usuario & "','')"
Call ConectionExecute(strSQL)
 
Call Bitacora("Registra", "NC Ope:" & gFondos.Operadora & " Plan:" & Trim(gFondos.Plan) & " Cont:" & gFondos.Contrato & " Monto:" & Trim(txtAporte))


'Aplicar a las sub Cuentas si Existe
If txtAporte.Locked Then
 For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 4
    If CCur(vGrid.Text) > 0 Then
       curSobrante = CCur(vGrid.Text)
       
        vGrid.col = 1
        strSQL = "Insert into fnd_SubCuentas_detalle(Idx,Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon) Values(" & vGrid.Text & "," & gFondos.Operadora & ",'" _
               & gFondos.Plan & "'," & gFondos.Contrato & ",dbo.MyGetdate()," & curSobrante * -1 & "," _
               & vProceso & ",'" & vTcon & "','" & vRecibo & "')"
        Call ConectionExecute(strSQL)
       
       vGrid.col = 5
       vAporteActual = CCur(vGrid.Text)
       vGrid.col = 6
       vRendiActual = CCur(vGrid.Text)
        
        If curSobrante >= vAporteActual Then
           vAplAporte = vAporteActual
           curSobrante = curSobrante - vAporteActual
        Else
           vAplAporte = curSobrante
           curSobrante = 0
        End If
         
        If curSobrante >= vRendiActual Then
           vAplRendi = vRendiActual
           curSobrante = curSobrante - vRendiActual
        Else
           vAplRendi = curSobrante
           curSobrante = 0
        End If
        
        vGrid.col = 1
        strSQL = "Update Fnd_subCuentas set Aportes = Aportes - " & vAplAporte _
               & ",rendimiento = rendimiento - " & vAplRendi _
               & " where cod_operadora=" & gFondos.Operadora & " and cod_plan = '" & gFondos.Plan _
               & "' and cod_contrato = " & gFondos.Contrato & " and IdX = " & vGrid.Text
        Call ConectionExecute(strSQL)
       
    End If
 Next i
End If
 
 
glogon.Conection.CommitTrans

Me.MousePointer = vbDefault
If GLOBALES.SysDocVersion = 1 Then
    Call sbgFNDImprimeRecibo(vRecibo, vTipoDoc, gFondos.Operadora)
Else
   Call sbImprimeRecibo(vRecibo, vTipoDoc)
End If

MsgBox "Anulación aplicada, con Nota de Crédito # " & vRecibo, vbInformation
Call sbLimpiaPantalla

Exit Sub

vError:
  Me.MousePointer = vbDefault
  Resume
  glogon.Conection.RollbackTrans
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion
End Sub

Private Sub Form_Load()
'Me.Icon = frmFNDOperadoras.Icon

vModulo = 18 'Fondo de Inversion
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call sbLimpiaPantalla
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaPantalla()
  
Dim strSQL As String, rs As New ADODB.Recordset

txtAporte = 0
txtAporte.Tag = "S"

vGrid.MaxRows = 0

strSQL = "select C.cedula,S.nombre,P.descripcion as PlanX,O.descripcion as OperadoraX" _
       & ",C.cod_plan,C.cod_contrato,C.cod_operadora,C.estado,C.fecha_Inicio,isnull(P.cuenta_Maestra,0) as CuentaMaestra" _
       & ",P.Tipo_CDP,P.permite_mov_cajas" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " inner join fnd_planes P on C.cod_plan = P.cod_plan and C.cod_operadora = P.cod_operadora" _
       & " inner join fnd_operadoras O on C.cod_operadora = O.cod_operadora" _
       & " where C.cod_operadora = " & gFondos.Operadora & " and C.cod_plan = '" & gFondos.Plan _
       & "' and C.cod_Contrato = " & gFondos.Contrato
Call OpenRecordSet(rs, strSQL)
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 txtContrato = rs!COD_CONTRATO
 txtOperadora = rs!operadoraX
 txtDescripcion = rs!PlanX
 gFondos.Contrato = rs!COD_CONTRATO
 txtFecha = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
 txtEstado.Tag = rs!Estado
 txtEstado = IIf((rs!Estado = "A"), "Activo", "Liquidado")
 
 If rs!tipo_cdp = 1 Then
    txtAporte.Tag = "N"
 End If
 
 If rs!PERMITE_MOV_CAJAS = 0 Then
    txtAporte.Tag = "N"
 End If

 If rs!CuentaMaestra = 1 Then
    txtAporte.Locked = True
    strSQL = "select IDx,Cedula,Nombre,0,aportes,rendimiento from fnd_subCuentas where cod_operadora = " & gFondos.Operadora _
           & " and cod_plan = '" & gFondos.Plan & "' and cod_contrato = " & gFondos.Contrato _
           & " and estado = 'A'"
   Call sbCargaGrid(vGrid, 6, strSQL)
   vGrid.MaxRows = vGrid.MaxRows - 1
 Else
    txtAporte.Locked = False
 End If
rs.Close

End Sub

Private Sub txtAporte_GotFocus()
On Error GoTo vError
 txtAporte = CCur(txtAporte)
vError:
End Sub

Private Sub txtAporte_LostFocus()
On Error GoTo vError
 txtAporte = Format(CCur(txtAporte), "Standard")
vError:
End Sub


