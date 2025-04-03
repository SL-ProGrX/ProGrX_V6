VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCntX_ConCierre 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consolidación - Actualización"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgLsw 
      Left            =   7080
      Top             =   2760
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
            Picture         =   "frmCntX_ConCierre.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConCierre.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConCierre.frx":020B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_ConCierre.frx":032B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2055
      Left            =   1200
      TabIndex        =   8
      Top             =   2160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLsw"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.TextBox txtPeriodo 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "(F4) Descripción del Periodo"
      Top             =   1560
      Width           =   4185
   End
   Begin VB.TextBox txtAnio 
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
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   4
      ToolTipText     =   "(F4) Año del Periodo"
      Top             =   1560
      Width           =   525
   End
   Begin VB.TextBox txtMes 
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
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      ToolTipText     =   "(F4) Mes del Periodo"
      Top             =   1560
      Width           =   315
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox cbo 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
   End
   Begin XtremeSuiteControls.PushButton cmdConCierre 
      Height          =   732
      Left            =   6360
      TabIndex        =   10
      Top             =   1200
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmCntX_ConCierre.frx":0459
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   7920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Proceso..:"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   8040
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Consolidación"
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
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actualización de la Contabilidad de Consolidación ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7812
   End
End
Attribute VB_Name = "frmCntX_ConCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbListaProgress(pItem As String, pIcono As Integer, pNota As String)
Dim i As Integer


For i = 1 To lsw.ListItems.Count
 If lsw.ListItems(i).Key = pItem Then
   lsw.ListItems(i).SmallIcon = pIcono
   lsw.ListItems(i).SubItems(1) = pNota
   Exit Sub
 End If
Next i


End Sub


Private Sub sbActualiza(pConsolida As Long, pMes As Integer, pAnio As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vContaBase As Long, vNivel As Integer
Dim vCon As New ADODB.Connection, rsX As New ADODB.Recordset
Dim vCadena As String, rsTmp As New ADODB.Recordset

'1. Verificar que el periodo no haya sido ya cargado anteriormente
'2. Verificar que todas las contabilidades miembro esten cerradas en el periodo indicado
'3. Totalizar los movimientos de cuentas
'4. Abrir el Periodo Consolidado para modificaciones

On Error GoTo vError

'vTran = False


' Set itmX = lsw.ListItems.Add(, "0x01", "Verificación de Contabilidades", , 4)
' Set itmX = lsw.ListItems.Add(, "0x02", "Verificación de Portales", , 4)
' Set itmX = lsw.ListItems.Add(, "0x03", "Comparando Catálogos", , 4)
' Set itmX = lsw.ListItems.Add(, "0x04", "Inicializando Contabilidad Base", , 4)
' Set itmX = lsw.ListItems.Add(, "0x05", "Consolidando Contabilidades Internas", , 4)
' Set itmX = lsw.ListItems.Add(, "0x06", "Consolidando Contabilidades Externas", , 4)
' Set itmX = lsw.ListItems.Add(, "0x07", "Aplicando Asientos Consolidados", , 4)



'Carga Parametros Iniciales
lbl.Caption = "Cargando Parámetros de Aplicación"
lbl.Refresh

strSQL = "select COD_CONTABILIDAD,nivel from CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & pConsolida
Call OpenRecordSet(rs, strSQL, 0)
    vContaBase = rs!COD_CONTABILIDAD
    vNivel = rs!nivel
rs.Close


'Verificaciones
lbl.Caption = "Verificacion Cierres de Periodo en Contabilidades Locales"
lbl.Refresh

DoEvents
strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where estado = 'P' and mes = " & pMes _
       & " and anio = " & pAnio & " and cod_contabilidad = " & vContaBase
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe > 0 Then
  rs.Close
  MsgBox "Este Periodo ya fue cerrado en la contabilidad base de la Consolidación...", vbInformation
  Exit Sub
End If
rs.Close

'Verifica Cierres de Periodos de Contabilidades Locales
strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where mes = " & pMes _
       & " and anio = " & pAnio & " and estado = 'P' and COD_CONTABILIDAD in(" _
       & "select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION_DET where cod_consolida = " & pConsolida & ")"
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe > 0 Then
  rs.Close
  MsgBox "Hay Contabilidades Locales de Este consolidado que NO han cerrado este periodo...", vbInformation
  Exit Sub
End If
rs.Close


lbl.Caption = "Verificacion Cierres de Periodo en Contabilidades Externas (Portal)"
lbl.Refresh

'Revisar Periodos de las contabilidades de externas
strSQL = "select C.cod_portal,C.COD_CONTABILIDAD,P.por_user,P.por_password,P.por_server,P.por_database" _
       & " from CNTX_CONSOLIDA_PORTALES_CON C inner join CNTX_CNTX_CONSOLIDA_PORTALES_CON P on C.cod_portal = P.cod_portal" _
       & " where C.cod_consolida = " & pConsolida
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 strSQL = fxPortalPrueba(Trim(rs!por_user), fxPortalCifrado(rs!por_password, "D"), Trim(rs!por_server), Trim(rs!por_database))
 If Len(strSQL) = 0 Then
   MsgBox "No se pudo realizar conección con el Portal : " & rs!cod_portal, vbExclamation
   rs.Close
   Exit Sub
 Else
   vCon.Open strSQL
   strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where mes = " & pMes _
          & " and anio = " & pAnio & " and estado = 'P' and COD_CONTABILIDAD = " & rs!COD_CONTABILIDAD
   rsX.Open strSQL, vCon, adOpenStatic
    If rsX!Existe > 0 Then
      MsgBox "El Periodo no se ha cerrado, en la contabilidad # " & rs!COD_CONTABILIDAD & " del Portal # " & rs!cod_portal, vbExclamation
      rsX.Close
      Exit Sub
    End If
   rsX.Close
   vCon.Close
 End If
 rs.MoveNext
Loop
rs.Close


'-----------------------------------------------------------------------------------------------------------------------------
'Proceso
Me.MousePointer = vbHourglass


lbl.Caption = "Iniciando Consolidacion Contabilidades Locales"
lbl.Refresh


'Crea Periodo Consolidado (verificar que no exista el periodo
strSQL = "select isnull(count(*),0) as Existe from CntX_Periodos where Anio = " & pAnio _
       & " and mes = " & pMes & " and cod_contabilidad = " & vContaBase
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe = 0 Then
    strSQL = "insert CntX_Periodos(anio,mes,estado,cod_contabilidad) values(" & pAnio _
           & "," & pMes & ",'P'," & vContaBase & ")"
    Call ConectionExecute(strSQL, 0)
End If

'Borra Balances Anteriores

'Poner asientos como no mayorizados y aplicarlos despues de cargar los balances de las contabilidades



'Inserta Movimientos de las Contabilidades locales
strSQL = "insert into con_movimientos(cod_consolida,COD_CONTABILIDAD,anio,mes,cod_cuenta,saldo_inicial" _
       & ",total_debitos,total_creditos) (select " & pConsolida & "," & vContaBase & "," & pAnio & "," _
       & pMes & ",M.cod_cuenta,isnull(sum(M.saldo_inicial),0) as SI," _
       & "isnull(sum(M.total_debitos),0) as TD,isnull(sum(M.Total_creditos),0) as TC" _
       & " from movimiento_cuentas M inner join Cuentas X on M.COD_CONTABILIDAD = X.COD_CONTABILIDAD" _
       & " and M.cod_cuenta = X.cod_cuenta" _
       & " inner join CNTX_CONSOLIDA_DEFINICION_DET C on M.COD_CONTABILIDAD = C.COD_CONTABILIDAD" _
       & " Where M.mes = " & pMes & " and M.anio = " & pAnio & " and C.cod_consolida = " _
       & pConsolida & " and X.nivel <= " & vNivel _
       & " group by M.cod_cuenta)"
Call ConectionExecute(strSQL, 0)
    
glogon.Conection.CommitTrans


lbl.Caption = "Iniciando Consolidacion Contabilidades Externas (Portales)"
lbl.Refresh

prgBar.Visible = True


'Agrupa por Portales las CONTABILIDADES y luego las agrega a los movimientos consolidado
'Revisar Periodos de las contabilidades de externas
strSQL = "select C.cod_portal,P.por_user,P.por_password,P.por_server,P.por_database" _
       & " from CNTX_CONSOLIDA_PORTALES_CON C inner join CNTX_CONSOLIDA_PORTALES_CON P on C.cod_portal = P.cod_portal" _
       & " where C.cod_consolida = " & pConsolida _
       & " group by C.cod_portal,P.por_user,P.por_password,P.por_server,P.por_database"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 strSQL = fxPortalPrueba(Trim(rs!por_user), fxPortalCifrado(rs!por_password, "D"), Trim(rs!por_server), Trim(rs!por_database))
 vCon.Open strSQL
 'Sacar listado de contabilidades x portal
 strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_PORTALES_CON where cod_consolida = " & pConsolida _
        & " and cod_portal = " & rs!cod_portal
 Call OpenRecordSet(rsX, strSQL, 0)
 vCadena = ""
 Do While Not rsX.EOF
   vCadena = vCadena & rsX!COD_CONTABILIDAD & ","
   rsX.MoveNext
 Loop
 rsX.Close
 
 If vCadena = "" Then
    vCadena = "0"
 Else
    vCadena = Mid(vCadena, 1, Len(vCadena) - 1)
 End If
 
 'Seleccionar Movimientos
 strSQL = "select M.cod_cuenta,isnull(sum(M.saldo_inicial),0) as SI," _
       & "isnull(sum(M.total_debitos),0) as TD,isnull(sum(M.Total_creditos),0) as TC" _
       & " from movimiento_cuentas M inner join Cuentas X on M.COD_CONTABILIDAD = X.COD_CONTABILIDAD" _
       & " and M.cod_cuenta = X.cod_cuenta" _
       & " Where M.mes = " & pMes & " and M.anio = " & pAnio & " and X.nivel <= " _
       & vNivel & " and X.COD_CONTABILIDAD in(" & vCadena & ")" _
       & " group by M.cod_cuenta"
 rsX.Open strSQL, vCon, adOpenStatic
 
 prgBar.Max = rsX.RecordCount + 2
 prgBar.Value = 1
 
 lbl.Caption = "Procesando Contabilidades del Portal # " & rs!cod_portal
 lbl.Refresh
 
 Do While Not rsX.EOF
 'Guardar Registro / Preguntar si existe
  strSQL = "select isnull(count(*),0) as Existe from con_movimientos where cod_cuenta = '" _
         & Trim(rsX!cod_cuenta) & "' and cod_consolida = " & pConsolida
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  If rsTmp!Existe = 0 Then
    'Insertar
    strSQL = "insert into con_movimientos(cod_consolida,COD_CONTABILIDAD,anio,mes,cod_cuenta,saldo_inicial" _
           & ",total_debitos,total_creditos) values( " & pConsolida & "," & vContaBase & "," & pAnio _
           & "," & pMes & ",'" & Trim(rsX!cod_cuenta) & "'," & rsX!si & "," & rsX!TD & "," & rsX!TC & ")"
  Else
    'Actualizar
    strSQL = "update con_movimientos set saldo_inicial = saldo_inicial + " & rsX!si _
           & ",total_debitos = total_debitos + " & rsX!TD _
           & ",total_creditos = total_creditos + " & rsX!TC _
           & " where cod_consolida = " & pConsolida & " and COD_CONTABILIDAD = " & vContaBase _
           & " and Anio = " & pAnio & " and mes = " & pMes _
           & " and cod_cuenta = '" & Trim(rsX!cod_cuenta) & "'"
    
  End If
  rsTmp.Close
  'Guarda Linea
  Call ConectionExecute(strSQL, 0)
  
  prgBar.Value = prgBar.Value + 1
  
  rsX.MoveNext
 Loop
 rsX.Close
 
 vCon.Close
 rs.MoveNext
Loop
rs.Close

'Mayorizar Asientos en la Contabilidad Base.

lbl.Caption = "Proceso Finalizado Satisfactoriamente"
prgBar.Visible = False

MsgBox "Periodo Consolidado Cargado Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cmdConCierre_Click()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta seguro que desea Actualizar la Consolidacion para este periodo...", vbYesNo)
If i = vbYes Then
    Me.MousePointer = vbHourglass
    Call sbActualiza(cbo.ItemData(cbo.ListIndex), txtMes, txtAnio)
    Me.MousePointer = vbDefault
End If
 
vError:
Me.MousePointer = vbDefault

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean
Dim itmX As ListItem
 
 vPaso = False
 
 strSQL = "select * from CNTX_CONSOLIDA_DEFINICION"
 Call OpenRecordSet(rs, strSQL, 0)
 cbo.Clear
 
 Do While Not rs.EOF
   cbo.AddItem Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
   cbo.ItemData(cbo.NewIndex) = rs!COD_CONSOLIDA
   rs.MoveNext
 Loop
 
 If vPaso Then
   rs.MoveFirst
   cbo.Text = Format(rs!COD_CONSOLIDA, "000") & " - " & rs!Descripcion
 End If
 rs.Close

 vPaso = True

 txtMes = Month(fxFechaServidor)
 txtAnio = Year(fxFechaServidor)
 
 
 'Lista de Aplicacion
 Set itmX = lsw.ListItems.Add(, "0x01", "Verificación de Contabilidades", , 4)
 Set itmX = lsw.ListItems.Add(, "0x02", "Verificación de Portales", , 4)
 Set itmX = lsw.ListItems.Add(, "0x03", "Comparando Catálogos", , 4)
 Set itmX = lsw.ListItems.Add(, "0x04", "Inicializando Contabilidad Base", , 4)
 Set itmX = lsw.ListItems.Add(, "0x05", "Consolidando Contabilidades Internas", , 4)
 Set itmX = lsw.ListItems.Add(, "0x06", "Consolidando Contabilidades Externas", , 4)
 Set itmX = lsw.ListItems.Add(, "0x07", "Aplicando Asientos Consolidados", , 4)
 
 

End Sub

Private Sub Form_Load()

 Call sbInicializa
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub


Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdConCierre.SetFocus
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub


