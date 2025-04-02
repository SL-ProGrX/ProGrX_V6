VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAF_CD_Liquidacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Liquidaciones Comites y Delegados"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7515
   Icon            =   "FrmAF_CD_ControlFactura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMontoDep 
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
      Left            =   5850
      TabIndex        =   18
      Top             =   1680
      Width           =   1515
   End
   Begin VB.TextBox txtRecibo 
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
      Left            =   5850
      TabIndex        =   2
      Top             =   1305
      Width           =   1515
   End
   Begin VB.TextBox txtComite 
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
      Left            =   2490
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
   Begin MSComctlLib.ListView lswOperaciones 
      Height          =   2565
      Left            =   180
      TabIndex        =   8
      Top             =   3435
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Op"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plan de Trabajo"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Monto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Solicitud"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Liquidación"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.TextBox txtMonto 
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
      Left            =   2490
      TabIndex        =   1
      Top             =   1680
      Width           =   1515
   End
   Begin VB.TextBox txtDescripcion 
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
      Height          =   840
      Left            =   2490
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2085
      Width           =   4875
   End
   Begin MSComctlLib.Toolbar tlbConsulta 
      Height          =   360
      Left            =   4155
      TabIndex        =   17
      Top             =   120
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   635
      ButtonWidth     =   1799
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Limpiar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ControlFactura.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ControlFactura.frx":1525B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAF_CD_ControlFactura.frx":15375
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Diferencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4125
      TabIndex        =   16
      Top             =   1755
      Width           =   1500
   End
   Begin VB.Label Label6 
      Caption         =   "No. Recibo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4125
      TabIndex        =   15
      Top             =   1365
      Width           =   900
   End
   Begin VB.Label Label11 
      Caption         =   "Up Comité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label13 
      Caption         =   "Comité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   13
      Top             =   615
      Width           =   660
   End
   Begin VB.Label Label15 
      Caption         =   "Delegado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   12
      Top             =   975
      Width           =   720
   End
   Begin VB.Label lblComite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2490
      TabIndex        =   11
      Top             =   540
      Width           =   4875
   End
   Begin VB.Label lblDelegado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2490
      TabIndex        =   10
      Top             =   915
      Width           =   4875
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Operaciones a Liquidar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   165
      TabIndex        =   9
      Top             =   3120
      Width           =   2700
   End
   Begin VB.Label lblGiro 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2490
      TabIndex        =   7
      Top             =   1290
      Width           =   1515
   End
   Begin VB.Label Label7 
      Caption         =   "Monto Girado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      TabIndex        =   6
      Top             =   1350
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "Monto Total en Facturas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   5
      Top             =   1710
      Width           =   2220
   End
   Begin VB.Label Label3 
      Caption         =   "Descripción de Liquidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   2085
      Width           =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7770
      Y1              =   3075
      Y2              =   3075
   End
End
Attribute VB_Name = "frmAF_CD_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim Z As Integer
Dim vAct As String
Dim vVerifica As Currency

Function FxNomComite(vUnidad As String)
   
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select descripcion from uprogramatica where codigo = '" & vUnidad & "'"
            rs.Open strSQL, glogon.Conection, adOpenStatic
   If rs.EOF Then
      FxNomComite = "No existe unidad definida "
   Else
      FxNomComite = rs!Descripcion
   End If

End Function


Private Function fxConsecutivo() As Long
Dim strSQL As String, rs As New ADODB.Recordset

  strSQL = "Select coalesce(Max(id_liquidacion),0) as consecutivo from afi_cd_controliquida"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  fxConsecutivo = rs!consecutivo + 1
  rs.Close

End Function


Private Sub sbLimpiar()

lswOperaciones.ListItems.Clear
txtDescripcion.Text = ""
lblGiro.Caption = ""
txtMonto.Text = 0
lblComite.Caption = ""
lblDelegado.Caption = ""
txtRecibo.Text = 0
txtMontoDep.Text = 0

End Sub

Private Sub cmdAplicar_Click()

Dim S As Integer
Dim i As Integer

strSQL = ""

If txtMonto.Text = Empty Then strSQL = strSQL & " El monto de la liquidación no es correcto" & vbCrLf


If strSQL <> "" Then
   MsgBox strSQL, vbInformation, "Faltan Los Siguientes Datos:"
   txtDescripcion.SetFocus
   Exit Sub
End If
 
 S = MsgBox("Desea cerrar la liquidación", vbYesNo + vbInformation, "Información")
       
   If vbYes = S Then
       For i = 1 To lswOperaciones.ListItems.Count
          If lswOperaciones.ListItems.Item(i).Checked = True Then
           
           strSQL = "insert afi_cd_controliquida (id_liquidacion,noperacion,descliquida,fecha,user_reg,monto," _
                    & "liquida,recibo,montorecibo) " _
                    & "values(" & fxConsecutivo & "," & lswOperaciones.ListItems.Item(i) & "," _
                    & "'" & txtDescripcion.Text & "','" & Format(fxFechaServidor, "yyyymmdd") & "'," _
                    & "'" & glogon.Usuario & "'," & CCur(lswOperaciones.ListItems.Item(i).SubItems(2)) & ",1," _
                    & "" & txtRecibo.Text & "," & CCur(txtMontoDep.Text) & ")"
                    glogon.Conection.Execute strSQL
                    
    
           strSQL = "update afi_cd_cuentas " _
                    & "set liquida_fecha = '" & Format(fxFechaServidor, "yyyymmdd") & "',estado = 'L'," _
                    & "liquida_usuario ='" & glogon.Usuario & "' where noperacion = " & lswOperaciones.ListItems.Item(i) & ""
                    glogon.Conection.Execute strSQL
          
          End If
      Next i
    MsgBox "La Liquidación fue registrada satisfactoriamente", vbInformation, "Información"
    End If
 Call sbLimpiar
End Sub

'Private Sub Cmdimprimir_Click()
'
'Dim strSQL As String
'On Error GoTo vError
'
'With frmContenedor.Crt
' .Reset
' .WindowShowGroupTree = True
' .WindowShowPrintSetupBtn = True
' .WindowShowRefreshBtn = True
' .WindowShowSearchBtn = True
' .WindowState = crptMaximized
' .Connect = glogon.ConectRPT
'
''Select Case True
''
'''Case OptComite = True
''         .WindowTitle = "Reporte Liquidaciones Realizadas por los Comites"
''         .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_ControlLiquidacion.rpt")
''
''         .Formulas(0) = "fxTitulo= 'LIQUIDACIONES REALIZADAS POR EL COMITE'"
''          strSQL = "{AFI_CD_AGRUPACOMITES.ID_PRICOMITE} = '" & txtComite.Text & "' and "
''          strSQL = strSQL & "cdate({AFI_CD_CONTROLIQUIDA.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
''          strSQL = strSQL & ") to Date (" & Format(dtpFinal.Value, "yyyy,mm,dd") & ")"
''         .SelectionFormula = strSQL
'
'
''Case OptTodosCom = True
''         .WindowTitle = "Reporte Liquidaciones Realizadas por los Comites"
''         .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_Cd_ControlLiquidacion.rpt")
''         .Formulas(0) = "fxTitulo= 'LIQUIDACIONES REALIZADAS POR LOS COMITES'"
''          strSQL = strSQL & "cdate({AFI_CD_CONTROLIQUIDA.FECHA}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
''          strSQL = strSQL & ") to Date (" & Format(dtpFinal.Value, "yyyy,mm,dd") & ")"
''         .SelectionFormula = strSQL
''
''End Select
''
'  .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
'  .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
'  .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
'  .Formulas(4) = "fxfecInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
'  .Formulas(5) = "fxfecfinal = '" & Format(dtpFinal.Value, "dd/mm/yyyy") & "'"
'  .Formulas(6) = "fxComite = '" & txtComite.Text & "'"
'  .PrintReport
'
'End With
'
'
'Exit Sub
'vError:
'  MsgBox Err.Description
'End Sub
'
Private Sub cmdLimpiar_Click()
Call sbLimpiar
End Sub

Private Sub Form_Load()
 'dtpInicio.Value = fxFechaServidor
 'dtpFinal.Value = dtpInicio.Value
 If GLOBALES.gTag <> Empty Then
   txtComite = GLOBALES.gTag
   Call TxtComite_KeyPress(vbKeyReturn)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   GLOBALES.gTag = Empty
End Sub

Private Sub LswOperaciones_Click()

Dim i As Integer
Dim vSuma As Currency
lblGiro.Caption = 0
txtMonto.Text = 0
txtMontoDep.Text = 0

For i = 1 To lswOperaciones.ListItems.Count
  If lswOperaciones.ListItems.Item(i).Checked = True Then
  vSuma = lswOperaciones.ListItems.Item(i).SubItems(2) + vSuma
  lblGiro.Caption = Format(vSuma, "Standard")
End If
Next i

End Sub

Private Sub tlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim S As Integer
Dim i As Integer


Select Case UCase(Button.Key)
  Case "NUEVO"
    Call sbLimpiar
  Case "APLICAR"

   strSQL = ""
    
    If txtMonto.Text = Empty Or CCur(CCur(txtMonto.Text) + CCur(txtMontoDep.Text)) < CCur(lblGiro.Caption) Then strSQL = strSQL & " El monto de la liquidación no es correcto" & vbCrLf


        If strSQL <> "" Then
           MsgBox strSQL, vbInformation, "Faltan Los Siguientes Datos:"
           txtDescripcion.SetFocus
           Exit Sub
        End If
 
   S = MsgBox("Desea cerrar la liquidación", vbYesNo + vbInformation, "Información")
       
   If vbYes = S Then
       For i = 1 To lswOperaciones.ListItems.Count
          If lswOperaciones.ListItems.Item(i).Checked = True Then
           
           strSQL = "insert afi_cd_controliquida (id_liquidacion,noperacion,descliquida,fecha,user_reg,monto," _
                    & "liquida,recibo,montorecibo) " _
                    & "values(" & fxConsecutivo & "," & lswOperaciones.ListItems.Item(i) & "," _
                    & "'" & txtDescripcion.Text & "','" & Format(fxFechaServidor, "yyyymmdd") & "'," _
                    & "'" & glogon.Usuario & "'," & CCur(lswOperaciones.ListItems.Item(i).SubItems(2)) & ",1," _
                    & "'" & txtRecibo.Text & "'," & CCur(txtMontoDep.Text) & ")"
                    glogon.Conection.Execute strSQL
                    
    
           strSQL = "update afi_cd_cuentas " _
                    & "set liquida_fecha = '" & Format(fxFechaServidor, "yyyymmdd") & "',estado = 'L'," _
                    & "liquida_usuario ='" & glogon.Usuario & "' where noperacion = " & lswOperaciones.ListItems.Item(i) & ""
                    glogon.Conection.Execute strSQL
          
          End If
      Next i
    MsgBox "La Liquidación fue registrada satisfactoriamente", vbInformation, "Información"
    End If
    Call sbLimpiar

  Case "REPORTE"
   strSQL = ""
   With frmContenedor.Crt
      .Reset
      .WindowShowGroupTree = True
      .WindowShowPrintSetupBtn = True
      .WindowShowRefreshBtn = True
      .WindowShowSearchBtn = True
      .WindowState = crptMaximized
      .Connect = glogon.ConectRPT
      
      .WindowTitle = "Reporte consulta de movimiento de actividades"
      .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_ControlLiquidacionEspecifico.rpt")
      .Formulas(0) = "fxTitulo= 'CONTROL DE LIQUIDACIONES POR COMITE'"
      strSQL = "({afi_cd_cuentas.cod_comite}) = '" & txtComite.Text & "'"
       'strSQL = strSQL & "cdate({vista_afi_cd_cuentasactividades.tesoreria_fecha}) "
       'in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
       'strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
       '.Formulas(4) = "fxFechaInicio = '" & Format(dtpInicio.Value, "dd/mm/yyyy") & "'"
       '.Formulas(5) = "fxFechaFinal = '" & Format(dtpCorte.Value, "dd/mm/yyyy") & "'"
      .SelectionFormula = strSQL
      .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .Formulas(3) = "fxUsuario='USER: " & glogon.Usuario & "'"
      
      .PrintReport
   End With
End Select



End Sub

Private Sub TxtComite_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case 13
     'LblComite.Caption = FxNomComite(TxtComite.Text)
     Call sbLiquidaComite
  Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtMonto_Change()
 
On Error GoTo vError
 
If txtMonto.Text = "" Then txtMonto.Text = 0

Exit Sub

vError:
 MsgBox "Error digitando el Monto", vbCritical, "Información"

End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
    
   tlbConsulta.Buttons.Item(2).Enabled = True
    
    Select Case KeyAscii
      Case 48 To 57, 8, 46, 44
      Case vbKeyReturn
        
        txtMontoDep.Text = Format((CCur(lblGiro.Caption) - CCur(txtMonto.Text)), "standard")
             
        If CCur(txtMonto.Text) > CCur(lblGiro.Caption) Then
            MsgBox "El monto a liquidar es mayor al monto que se giro", vbCritical, "Información"
        ElseIf CCur(txtMonto.Text) < CCur(lblGiro.Caption) Then
            MsgBox "Se tiene una diferencia de " & Format((CCur(lblGiro.Caption) - CCur(txtMonto.Text)), "Currency") & _
                 ", por depositar", vbInformation
           
           tlbConsulta.Buttons.Item(2).Enabled = False
          
        End If
        
        txtMonto.Text = Format(txtMonto.Text, "Standard")
        
      Case Else
        KeyAscii = 0
    End Select
End Sub



Private Sub sbLiquidaComite()
Dim itmX As ListItem

lswOperaciones.ListItems.Clear

If txtComite.Text = "" Then
   MsgBox "No Ingreso la Unidad Programática del Comité", vbInformation, "Información"
   Exit Sub
End If


strSQL = "select C.noperacion,C.cod_comite,U.descripcion,L.descliquida,L.recibo,L.montorecibo, " _
          & "C.cedula,S.nombre,C.liquida_usuario,C.liquida_fecha,L.fecha,L.liquida,L.monto," _
          & "C.tesoreria_nsolicitud from uprogramatica U inner join afi_cd_cuentas C " _
          & "on U.codigo = C.cod_comite left join afi_cd_controliquida L on C.noperacion = L.noperacion " _
          & "inner join socios S on C.cedula = S.cedula " _
          & "Where C.cod_comite = '" & txtComite.Text & "'"
          rs.Open strSQL, glogon.Conection, adOpenStatic
   
  
  If Not rs.EOF Then
        
        
        txtDescripcion.Text = IIf(Not IsNull(rs!descliquida), rs!descliquida, "Sin Descripcion")
        lblComite.Caption = IIf(Not IsNull(rs!Descripcion), rs!Descripcion, "Sin Descripcion")
        lblDelegado.Caption = IIf(Not IsNull(rs!Nombre), rs!Nombre, "Sin Descripcion")
        
        
        End If
rs.Close

strSQL = "select A.noperacion,C.notas,sum(A.monto)as Monto,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha " _
         & "from afi_cd_cuentas C inner join  afi_cd_cuentas_actividades A " _
         & "on C.noperacion = A.noperacion " _
         & "where C.cod_comite = '" & txtComite.Text & "' and estado ='T' " _
         & "group by C.notas,A.noperacion,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha"
         rs.Open strSQL, glogon.Conection, adOpenStatic

While Not rs.EOF
      Set itmX = lswOperaciones.ListItems.Add(, , Trim(rs!Noperacion))
      itmX.SubItems(1) = rs!notas
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!TESORERIA_NSOLICITUD
      itmX.SubItems(4) = Format(rs!liquida_fecha, "dd/mm/yyyy")
      rs.MoveNext
   Wend
rs.Close

End Sub

Private Sub txtMonto_LostFocus()
 If CCur(txtMonto.Text + txtMontoDep.Text) = CCur(lblGiro.Caption) Then
    tlbConsulta.Buttons.Item(2).Enabled = False
 End If
End Sub

Private Sub txtMontoDep_LostFocus()
   If CCur(txtMonto.Text + txtMontoDep.Text) = CCur(lblGiro.Caption) Then
     tlbConsulta.Buttons.Item(2).Enabled = True
   Else
     tlbConsulta.Buttons.Item(2).Enabled = False
   End If
End Sub

Private Sub Txtrecibo_KeyPress(KeyAscii As Integer)
 
  Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        txtDescripcion.SetFocus
      Case Else
        KeyAscii = 0
    End Select

End Sub

Private Sub Txtrecibo_LostFocus()
  If txtRecibo.Text = "" Then txtRecibo.Text = 0
End Sub
