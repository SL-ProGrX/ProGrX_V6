VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFNDAcuerdoVivienda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acuerdo de Vivienda ASECCSS"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArchivo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   6855
   End
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   6900
   End
   Begin MSComDlg.CommonDialog Cmd 
      Left            =   0
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDAcuerdoVivienda.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDAcuerdoVivienda.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDAcuerdoVivienda.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDAcuerdoVivienda.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   660
      Left            =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   1164
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cargar"
            Object.ToolTipText     =   "Cargar información"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
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
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmFNDAcuerdoVivienda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function fxCedulaOp(pOperacion As Long) As String
Dim strSQL As String, rsX As New ADODB.Recordset

strSQL = "select cedula from reg_creditos where id_solicitud = " & pOperacion
rsX.Open strSQL, glogon.Conection, adOpenStatic

    fxCedulaOp = Trim(rsX!Cedula)

rsX.Close

End Function

Private Function fxFNDConsecutivoContrato(vOperadora As Integer, vPlan As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

fxFNDConsecutivoContrato = 0

strSQL = "Select Consecutivo From fnd_planes where cod_operadora=" & vOperadora _
       & " and cod_plan='" & vPlan & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
    fxFNDConsecutivoContrato = rs!Consecutivo + 1
    strSQL = "Update fnd_planes set Consecutivo = Consecutivo + 1 " _
           & " Where cod_operadora = " & vOperadora & " and cod_plan = '" & vPlan & "'"
    glogon.Conection.Execute strSQL
End If
rs.Close

End Function

Private Function fxContrato(pCedula As String, pPlan As String) As Long
Dim strSQL As String, rsX As New ADODB.Recordset
Dim vResultado As Long

strSQL = "select cod_contrato from fnd_contratos where cod_plan = '" & pPlan _
       & "' and cedula = '" & pCedula & "' and estado = 'A'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  'Crear el Contrato y devolver y numero
  vResultado = fxFNDConsecutivoContrato(1, pPlan)
  
  strSQL = "insert FND_Contratos(Cod_operadora,Cod_plan,Cod_Contrato,Cedula,Cod_Vendedor," _
         & "Estado,Fecha_Inicio,Plazo,Monto,Renueva,Inc_Anual,Inc_Tipo,Ind_comision,Ind_deduccion," _
         & "Cod_Banco,Cuenta_Ahorros,Tipo_Pago,CapExc,Aportes,Rendimiento) values(1,'" _
         & pPlan & "'," & vResultado & ",'" & Trim(pCedula) & "','101020','A',GetDate()," _
         & "12,0,'N',0,'M',0,0,1,'','CK',0,0,0)"
  glogon.Conection.Execute strSQL
Else
  vResultado = rsX!cod_contrato
End If
rsX.Close

fxContrato = vResultado

End Function



Private Sub sbCargaDeducciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curAporte1 As Currency, curAporte2 As Currency
Dim curRend1 As Currency, curRend2 As Currency, vOperadora As Integer


Dim TmpAporte1 As Currency, TmpAporte2 As Currency
Dim TmpRend1 As Currency, TmpRend2 As Currency

Dim vPlan As String, vCedula As String, vContrato As Long
Dim vND As Long, vFecha As Date, vProceso As Long
Dim vCuentaFnd As String, vCuentaRendi As String, vCuentaGasto As String
Dim vCuentaGst As String, vCuentaInt As String
Dim vTipoDoc As String, strLinea(4) As String
'Gasto de Periodos Anteriores 505-20-01-00
'Periodo Actual Intereses: 401-10-01-00

curAporte1 = 0
curAporte2 = 0
curRend1 = 0
curRend2 = 0
vOperadora = 1
vPlan = "AHVI"
vTipoDoc = "ND"

DaoControl.Connect = "Excel 8.0;"
DaoControl.DatabaseName = txtArchivo.Text
DaoControl.RecordSource = "Resultados$"
DaoControl.Refresh

vFecha = fxFechaServidor
vProceso = 201009

rs.Open "select cuenta_conta,cuenta_rendimiento,cuenta_gasto from fnd_planes where cod_operadora = " & vOperadora _
  & " and cod_plan = '" & vPlan & "'", glogon.Conection, adOpenStatic
    vCuentaFnd = Trim(rs!Cuenta_Conta)
    vCuentaRendi = Trim(rs!Cuenta_Rendimiento)
    vCuentaGasto = Trim(rs!Cuenta_Gasto)
rs.Close

vCuentaGst = "505200100"
vCuentaInt = "401100100"
vND = fxgFNDDocumentoConsecutivo("ND", "1")

With DaoControl.Recordset
    Do While Not .EOF
      vCedula = fxCedulaOp(!Operacion)
      vContrato = fxContrato(vCedula, vPlan)
                  
    TmpAporte1 = Round(!Aporte1, 2)
    TmpAporte2 = Round(!Aporte2, 2)
    TmpRend1 = Round(!Rendimiento1, 2)
    TmpRend2 = Round(!Rendimiento2, 2)
    
    strSQL = "Update FND_Contratos set Aportes = Aportes + " & TmpAporte1 + TmpAporte2 _
           & ",rendimiento = rendimiento + " & TmpRend1 + TmpRend2 _
           & " where cod_operadora = " & vOperadora & " and cod_plan = '" & vPlan _
           & "' and cod_contrato = " & vContrato
    glogon.Conection.Execute strSQL
    
    'Inserta Detalle
    If TmpAporte1 > 0 Then
        strSQL = "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,Fecha_Acredita) Values(" & vOperadora _
               & ",'" & vPlan & "'," & vContrato & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & TmpAporte1 & "," & vProceso _
               & ",'8','" & vND & ".Op." & !Operacion & ".Ap.08-09','" & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        curAporte1 = curAporte1 + TmpAporte1
    End If
            
    If TmpAporte2 > 0 Then
        strSQL = "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,Fecha_Acredita) Values(" & vOperadora _
               & ",'" & vPlan & "'," & vContrato & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & TmpAporte2 & "," & vProceso _
               & ",'8','" & vND & ".Op." & !Operacion & ".Ap.09-10','" & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        curAporte2 = curAporte2 + TmpAporte2
    End If
            
            
    If TmpRend1 > 0 Then
        strSQL = "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,Fecha_Acredita) Values(" & vOperadora _
               & ",'" & vPlan & "'," & vContrato & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & TmpRend1 & "," & vProceso _
               & ",'8','" & vND & ".Op." & !Operacion & ".Rnd.08-09','" & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        curRend1 = curRend1 + TmpRend1
    End If
            
    If TmpRend2 > 0 Then
        strSQL = "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato," _
               & "Fecha,Monto,Fecha_Proceso,Tcon,Ncon,Fecha_Acredita) Values(" & vOperadora _
               & ",'" & vPlan & "'," & vContrato & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & TmpRend2 & "," & vProceso _
               & ",'8','" & vND & ".Op." & !Operacion & ".Rnd.09-10','" & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        curRend2 = curRend2 + TmpRend2
    End If
      .MoveNext
    Loop
End With
        
'Asiento

strLinea(1) = "Aporte Per.Ant: " & Format(curAporte1, "Standard")
strLinea(2) = "Aporte Per.Act: " & Format(curAporte2, "Standard")
strLinea(3) = "Rendi. Per.Ant: " & Format(curRend1, "Standard")
strLinea(4) = "Rendi. Per.Act: " & Format(curRend2, "Standard")

strSQL = "insert fnd_documentos(tipo,id_documento,cod_operadora,cliente,concepto,fecha," _
        & "monto,usuario,detalle1,detalle2,detalle3,detalle4,detalle,dp)" _
        & " values('" & vTipoDoc & "'," & vND & "," _
        & vOperadora & ",'CLIENTE GENERAL','APLICACION ACUERDO JD VIVIENDA" _
        & "',getdate()," & curAporte1 + curAporte2 + curRend1 + curRend2 & ",'" & Trim(glogon.Usuario) _
        & "','" & Mid(strLinea(1), 1, 40) & "','" & Mid(strLinea(2), 1, 40) & "','" & Mid(strLinea(3), 1, 40) & "','" _
        & Mid(strLinea(4), 1, 40) & "','','')"
glogon.Conection.Execute strSQL


'Aportes Acreditados
strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaGst & "'," & curAporte1 & ",'D')"
glogon.Conection.Execute strSQL

strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaInt & "'," & curAporte2 & ",'D')"
glogon.Conection.Execute strSQL

strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaFnd & "'," & curAporte1 + curAporte2 & ",'C')"
glogon.Conection.Execute strSQL

'Rendimiento
strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaGst & "'," & curRend1 & ",'D')"
glogon.Conection.Execute strSQL

strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaGasto & "'," & curRend2 & ",'D')"
glogon.Conection.Execute strSQL


strSQL = "insert fnd_asientos(Cod_Operadora,Tipo,Id_documento,Fnd_Cuenta,Fnd_Monto," _
       & "Fnd_Debehaber) Values(" & vOperadora & ",'" & vTipoDoc & "'," _
       & vND & ",'" & vCuentaRendi & "'," & curRend1 + curRend2 & ",'C')"
glogon.Conection.Execute strSQL


Me.MousePointer = vbDefault
MsgBox "Casos Procesados y Actualizados"

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical
End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "buscar"
        
        txtArchivo.Text = ""
        
        With Cmd
                 .InitDir = "C:\"
                 .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL 97-2003]..."
                 .Filter = "*.xls"
                 .ShowOpen
                 
                 If .FileName = "" Then
                   MsgBox "Archivo no válido...", vbExclamation
                   Exit Sub
                 End If
                 
                 If UCase(Right(.FileName, 3)) <> "XLS" Then
                   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                   Exit Sub
                 End If
                 txtArchivo.Text = .FileName
        
        End With

  Case "cargar"
    Call sbCargaDeducciones
  
End Select

End Sub
