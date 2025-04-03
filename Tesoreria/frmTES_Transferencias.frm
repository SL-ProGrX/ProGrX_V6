VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmTES_Transferencias 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesando..."
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   648
      Left            =   5640
      TabIndex        =   4
      Top             =   3120
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1143
      _StockProps     =   79
      Caption         =   "Aceptar"
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
      Picture         =   "frmTES_Transferencias.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdReversar 
      Height          =   648
      Left            =   4320
      TabIndex        =   5
      Top             =   3120
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   1143
      _StockProps     =   79
      Caption         =   "Reversar"
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
      Picture         =   "frmTES_Transferencias.frx":07DE
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   612
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7452
      _Version        =   1441793
      _ExtentX        =   13144
      _ExtentY        =   1080
      _StockProps     =   14
      Caption         =   "Control de Emisión de Transferencias Electrónicas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label lblArchivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   852
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   7188
   End
   Begin VB.Label lblBanco 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   612
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7212
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Archivo :"
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
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1572
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cuenta :"
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
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1572
   End
End
Attribute VB_Name = "frmTES_Transferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim i As Long, vDocumento As String
Dim curMonto As Currency
Dim vFecha As Date


Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

i = 0

curMonto = 0

'rs.Open gstrQuery, glogon.Conection, adOpenStatic
Call OpenRecordSet(rs, gstrQuery)
Do While Not rs.EOF
   
   If i = 0 Then
      i = fxTesTipoDocConsecInterno(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "+", gTesGlobal.BancoPlan)
   Else
      i = i + 1
   End If
   
   vDocumento = Format(i, "0000")
      
   curMonto = curMonto + rs!Monto
      
   strSQL = strSQL & Space(10) & "Update Tes_Transacciones Set Estado='T' , Fecha_Emision = '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
          & "',Ubicacion_Actual = 'T',FECHA_TRASLADO = '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',NDocumento = '" _
          & vDocumento & "',user_genera = '" & glogon.Usuario & "',documento_base = '" & gTesGlobal.BancoConsec _
          & "', COD_PLAN = '" & gTesGlobal.BancoPlan & "'" _
          & " Where NSolicitud=" & rs!NSolicitud
   
    'Bitacora Especial
    strSQL = strSQL & Space(10) & "exec spTesBitacora " & rs!NSolicitud & ",'10','" & Mid("Transferencia...: " & gTesGlobal.BancoConsec, 1, 150) _
           & "','" & glogon.Usuario & "'"
   
    'Afecta Saldo en Bancos
    strSQL = strSQL & Space(10) & "exec spTESAfectaBancos " & rs!NSolicitud & ",'E'"
   
   'Actualiza Cuentas Corrientes
   If rs!Modulo = "CC" And rs!submodulo = "C" Then
    If IsNumeric(rs!Detalle1) Then
'       Call sbTESActualizaCC(rs!Codigo, rs!Tipo, vDocumento, rs!id_Banco, rs!Detalle1, rs!modulo, rs!submodulo, IIf(IsNull(rs!referencia), 0, rs!referencia))
        
        If IIf(IsNull(rs!Referencia), 0, rs!Referencia) > 0 Then
            'TIENE REFERENCIA
            strSQL = strSQL & Space(10) & "Update DesemBolsos Set Cod_Banco=" & rs!Id_Banco & ",TDocumento='" & rs!Tipo & "'," _
                   & "NDocumento='" & vDocumento & "' Where ID_Desembolso=" & Trim(rs!Codigo)
        Else
            'NO TIENE REFERENCIA
            strSQL = strSQL & Space(10) & "Update Reg_Creditos Set Cod_Banco = " & rs!Id_Banco & ",Documento_Referido = '" & rs!Tipo _
                   & "-" & vDocumento & "' Where ID_Solicitud=" & rs!Detalle1
        End If
    End If
   End If
   
   
   If Len(strSQL) > 20000 Then
       glogon.Conection.BeginTrans
            Call ConectionExecute(strSQL)
       glogon.Conection.CommitTrans

       strSQL = ""
   End If
   
   rs.MoveNext
Loop
rs.Close

'Procesa Lote Final
If Len(strSQL) > 0 Then
       glogon.Conection.BeginTrans
            Call ConectionExecute(strSQL)
       glogon.Conection.CommitTrans
End If

'Actualiza Consecutivo Interno
strSQL = "update tes_banco_docs set CONSECUTIVO_DET = " & i _
       & " where Tipo = '" & gTesGlobal.BancoTDoc & "' and id_banco = " & gTesGlobal.BancoID
Call ConectionExecute(strSQL)




Call sbTesReporteTransferencia(gTesGlobal.BancoID, gTesGlobal.BancoConsec, "C", gTesGlobal.BancoTDoc, gTesGlobal.BancoPlan)
Call sbTesReporteTransferencia(gTesGlobal.BancoID, gTesGlobal.BancoConsec, "D", gTesGlobal.BancoTDoc, gTesGlobal.BancoPlan)

Me.MousePointer = vbDefault

Me.Hide

Unload Me

End Sub

Private Sub cmdReversar_Click()
Dim strSQL As String, lngX As Long

Me.MousePointer = vbHourglass

Kill (lblArchivo)

lngX = fxTesTipoDocConsec(gTesGlobal.BancoID, gTesGlobal.BancoTDoc, "-", gTesGlobal.BancoPlan)


Me.MousePointer = vbDefault

Me.Hide

frmTES_EmisionDocumentos.Show

Unload Me

End Sub



Private Sub Form_Load()
 vModulo = 9

End Sub

