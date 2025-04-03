VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_DesemConceptos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Desembolsos/Retención de Crédito"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   Icon            =   "frmCR_DesemConceptos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   14115
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6012
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   13812
      _Version        =   524288
      _ExtentX        =   24363
      _ExtentY        =   10604
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
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_DesemConceptos.frx":08CA
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Conceptos para Desembolsos y/o Rebajos para Créditos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   14292
   End
End
Attribute VB_Name = "frmCR_DesemConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3

End Sub


Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""
fxVerifica = True

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

'If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique un código para este cargo" & vbCrLf


vGrid.col = 3
If Not fxCntX_CuentaValida(vGrid.Text) Then vMensaje = vMensaje & " - Especifique una cuenta contable válida!" & vbCrLf

vGrid.col = 6
If vGrid.Value = vbChecked Then
    vGrid.col = 7
    If Not fxCntX_CuentaValida(vGrid.Text) Then vMensaje = vMensaje & " - La Cuenta Contable para Diferir no es válida!" & vbCrLf
End If

If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxVerifica = False
End If

End Function


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

If Not fxVerifica Then Exit Function

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
If vGrid.Text = "" Then
    vGrid.col = 2
    strSQL = "insert CONCEPTO_DESEMB(descripcion,cod_cuenta,retiene,modifica,DIFIERE,DIFIERE_CUENTA,ACTIVO)" _
           & " values('" & vGrid.Text & "','"
    vGrid.col = 3
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Value & ","
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value & ","
    vGrid.col = 6
    strSQL = strSQL & vGrid.Value & ",'"
    vGrid.col = 7
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 8
    strSQL = strSQL & vGrid.Value & ")"
      
    Call ConectionExecute(strSQL)
  
    Call Bitacora("Registra", "Concepto Desembolso : " & vGrid.Text)
  
  
    vGrid.col = 2
    strSQL = "select max(COD_CONDEB) as ultimo from CONCEPTO_DESEMB where descripcion = '" & vGrid.Text & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
   
   Else 'Actualizar

    vGrid.col = 2
    strSQL = "update CONCEPTO_DESEMB set descripcion = '" & vGrid.Text _
           & "',COD_CUENTA = '"
    vGrid.col = 3
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "',RETIENE = "
    vGrid.col = 4
    strSQL = strSQL & vGrid.Value & ",MODIFICA = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value & ",DIFIERE = "
    vGrid.col = 6
    strSQL = strSQL & vGrid.Value & ",DIFIERE_CUENTA = '"
    vGrid.col = 7
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "',ACTIVO = "
    vGrid.col = 8
    strSQL = strSQL & vGrid.Value & " where COD_CONDEB = "
    vGrid.col = 1
    strSQL = strSQL & vGrid.Text
    Call ConectionExecute(strSQL)
    
   End If

   vGrid.col = 1
   fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If i > 0 Then
        vGrid.Text = i
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
  End If
End If



If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 3 Or vGrid.ActiveCol = 7) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.ActiveCol
  Call sbgCntCuentaConsulta
   vGrid.Text = fxCntX_CuentaFormato(True, gBusquedas.Resultado, 0)
End If

End Sub

Private Sub sbCargaGridLocal(pGrid As Object, pGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

pGrid.MaxCols = pGridMaxCol
pGrid.MaxRows = 1
pGrid.Row = pGrid.MaxRows
For i = 1 To pGrid.MaxCols
 pGrid.col = i
 pGrid.Text = ""
Next i

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  pGrid.Row = pGrid.MaxRows
  For i = 1 To pGrid.MaxCols
     pGrid.col = i
     Select Case i
         Case 3, 7
            pGrid.Text = fxgCntCuentaFormato(True, CStr(rs.Fields(i - 1).Value), 0)
         Case Else
            pGrid.Text = CStr(rs.Fields(i - 1).Value)
     End Select
  Next i
  pGrid.MaxRows = pGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

End Sub


Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Cargar en el Objeto adoData Todas las causas de renuncia que existen.
'REFERENCIAS:   sbToolBarIconos - (Carga los iconos para la barra de herramientas)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String

On Error GoTo error

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select COD_CONDEB,descripcion,cod_cuenta,retiene,modifica,difiere,DIFIERE_CUENTA,activo" _
       & " from CONCEPTO_DESEMB order by descripcion"
Call sbCargaGridLocal(vGrid, 8, strSQL)

Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Realizar el mantenimiento de las causas de renuncia.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de datos)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Integer, strSQL As String


On Error GoTo vError


    Select Case Button.Key
        Case "insertar"
             vGrid.MaxRows = vGrid.MaxRows + 1
        
        Case "borrar"
            i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
            If i = vbYes Then
               vGrid.Row = vGrid.ActiveRow
               vGrid.col = 1
               strSQL = "delete CONCEPTO_DESEMB where COD_CONDEB= " & vGrid.Text
               Call ConectionExecute(strSQL)
               
               strSQL = vGrid.Text
               vGrid.col = 2
               Call Bitacora("Elimina", "Concepto de Desembolso : " & strSQL & " - " & vGrid.Text)
               vGrid.col = 1
               
                strSQL = "select COD_CONDEB,descripcion,cod_cuenta,retiene,modifica,difiere,DIFIERE_CUENTA,activo" _
                       & " from CONCEPTO_DESEMB order by descripcion"
                Call sbCargaGridLocal(vGrid, 8, strSQL)
            End If
            
        Case "reportes"
           With frmContenedor.Crt
             .Reset
             .WindowShowPrintSetupBtn = True
             .WindowShowRefreshBtn = True
             .WindowShowSearchBtn = True
             .WindowState = crptMaximized
             .WindowTitle = "Reportes del Módulo de Crédito"
             
             .Connect = glogon.ConectRPT

             .ReportFileName = SIFGlobal.fxPathReportes("Credito_DesembolsoCNT.rpt") 'Ojo falta reporte
             .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
             .PrintReport
           End With
             
        Case "ayuda"
            frmContenedor.CD.HelpContext = Me.HelpContextID
            frmContenedor.CD.ShowHelp
           
    End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If col = 3 Or col = 7 Then
    vGrid.Row = Row
    vGrid.col = col
    vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text, 0)
End If

End Sub
