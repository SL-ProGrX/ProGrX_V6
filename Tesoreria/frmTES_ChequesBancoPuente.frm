VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_ChequesBancoPuente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cheques en Banco puente"
   ClientHeight    =   7032
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   10524
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7032
   ScaleWidth      =   10524
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.ComboBox cboBancos 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10212
      _Version        =   524288
      _ExtentX        =   18013
      _ExtentY        =   8700
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1000000
      ScrollBars      =   2
      SpreadDesigner  =   "frmTES_ChequesBancoPuente.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   672
      Left            =   8520
      TabIndex        =   5
      Top             =   6120
      Width           =   1692
      _Version        =   1245187
      _ExtentX        =   2984
      _ExtentY        =   1185
      _StockProps     =   79
      Caption         =   "&Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmTES_ChequesBancoPuente.frx":062F
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Trasladar al Banco"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bancos Puente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmTES_ChequesBancoPuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboBancos_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iCantidadFilas As Long
strSQL = "select 0,nsolicitud,codigo,beneficiario,monto,fecha_solicitud from cheques where id_banco = " & cboBancos.ItemData(cboBancos.ListIndex) & " and  ESTADO = 'P' and tipo = 'CK'"
Call sbCargaGrid(vGrid, 6, strSQL)
'Call OpenRecordSet(rs, strSQL)
'If Not rs.EOF Then vGrid.MaxRows = rs.RecordCount
'iCantidadFilas = 1
''vGrid.MaxRows = 0
'Do While Not rs.EOF
'
'   vGrid.Row = iCantidadFilas
'   vGrid.Col = 2
'   vGrid.Text = str(rs!nsolicitud)
'   vGrid.Col = 3
'   vGrid.Text = rs!codigo
'   vGrid.Col = 4
'   vGrid.Text = rs!beneficiario
'   vGrid.Col = 5
'   vGrid.Text = Format(rs!monto, "Standard")
'   vGrid.Col = 6
'   vGrid.Text = Format(rs!fecha_solicitud, "dd/mm/yyyy")
'   rs.MoveNext
'
'   iCantidadFilas = iCantidadFilas + 1
'
'Loop


End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuentaD As String, vCuentaI As String
Dim i As Integer
On Error GoTo vError
'1. Verificar que no sea el mismo banco
'2. Verificar que el asiento este pendiente de traspaso
'3. Cambiar la cuenta del banco
'4. Restabler Saldo en Bancos / Traslada Saldos Afectados

For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        vGrid.col = 1
    If vGrid.Value = 1 Then
   
        vGrid.col = 2
        If cboBancos.ItemData(cboBancos.ListIndex) = cbo.ItemData(cbo.ListIndex) Then Exit Sub
        
        vCuentaI = fxCuentaBanco(cboBancos.ItemData(cboBancos.ListIndex))
        vCuentaD = fxCuentaBanco(cbo.ItemData(cbo.ListIndex))
        
        
        strSQL = "select estado_asiento from cheques where nsolicitud = " & vGrid.Text
        Call OpenRecordSet(rs, strSQL)
        If rs!estado_asiento = "G" Then
          rs.Close
          MsgBox "El asiento de esta solicitud ya fue generado, no se puede reclasificar...", vbInformation
          Exit Sub
        End If
        rs.Close
        
        strSQL = "update cheques set id_banco = " & cbo.ItemData(cbo.ListIndex) _
               & " where nsolicitud = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = "update ck_detalle set cuenta_contable = '" & vCuentaD _
               & "' where cuenta_contable = '" & vCuentaI & "' and nsolicitud = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Modifica", "Cambia de Banco Solicitud N. " & Trim(vGrid.Text))

    End If
Next i
MsgBox "Cambio de Banco Realizado Satisfactoriamente...", vbInformation
Call cboBancos_Click
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 9

Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Genera")

strSQL = "select id_banco,descripcion from tes_bancos where estado = 'A' and puente  = 1"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   cboBancos.AddItem IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
   cboBancos.ItemData(cboBancos.NewIndex) = rs!id_Banco
   rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboBancos.Text = rs!Descripcion
End If
rs.Close


End Sub


Private Function fxCuentaBanco(vBanco As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select ctaconta as Cuenta from tes_bancos where id_banco = " & vBanco
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 fxCuentaBanco = Trim(rs!Cuenta)
Else
 fxCuentaBanco = ""
End If
rs.Close

End Function
