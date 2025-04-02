VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmPosCajaLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de Sesión en Cajas"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione su Caja y digite su contraseña"
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
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   5532
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPosCajaLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
Dim strSQL As String

i = 0

gCajas.Apertura = 0

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


txtUsuario = glogon.Usuario
txtClave = ""

strSQL = "select rtrim(cod_caja) as 'IdX' , rtrim(nombre) as 'ItmX'" _
       & " from pv_cajas where estado = 'A' and usuario = '" _
       & glogon.Usuario & "' order by cod_caja"

Call sbCbo_Llena_New(cbo, strSQL, False, True)

End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

'1. Verificar que el Estado este Activa / en el Load esta validado
'2. Que no se encuentre Bloqueada
'3. Verificar si la caja esta abierta (Apertura) y Sacar el Consecutivo
'   de la apertura.

If i > 3 And KeyCode = vbKeyReturn Then
  MsgBox "No se permiten más intentos...", vbExclamation
  Unload Me
End If



If KeyCode = vbKeyReturn Then
 i = i + 1
 strSQL = "select bloqueo from pv_cajas where usuario = '" _
        & txtUsuario & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "' and clave = '" _
        & fxPosEncrypta(txtClave) & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs.EOF And rs.BOF Then
   MsgBox "Caja: verifique su Usuario y Clave para Esta Caja ...", vbExclamation
 Else
  If rs!bloqueo = 0 Then
     gCajas.Caja = cbo.ItemData(cbo.ListIndex)
     gCajas.Usuario = txtUsuario
     
     'Consulta la caja para verificar que tenga una apertura existente
     rs.Close
     strSQL = "select cod_ac from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
            & "' and usuario = '" & gCajas.Usuario & "' and estado = 'A'"
     Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
        gCajas.Apertura = rs!cod_ac
        
        
        rs.Close
        'Inicializa Datos de Cajas
            strSQL = "select * from pv_cajas" _
                   & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
            Call OpenRecordSet(rs, strSQL)
                gCajas.Agente = rs!def_agente
                gCajas.Bodega = rs!def_bodega
                gCajas.Cliente = rs!def_cliente
                gCajas.Precio = rs!def_precio
                gCajas.Nombre = rs!Nombre
                gCajas.BodegaDesc = fxSIFCCodigos("D", rs!def_bodega, "bodegas")
                gCajas.ModFechas = IIf((rs!modifica_fechas = 1), True, False)
                gCajas.ModPrecios = IIf((rs!modifica_precio = 1), True, False)
                gCajas.VentaExenta = IIf((rs!venta_Exenta = 1), True, False)
                
                
                gCajas.Display = rs!POS_DISPLAY
                gCajas.Autorizador = rs!Autorizador & ""
                gCajas.F_Factura = rs!Formato_Factura
                gCajas.F_Especial = IIf((rs!FORMATO_ESPECIAL = 1), True, False)
                gCajas.F_Archivo = rs!FORMATO_ESPECIAL_ARCHIVO
            'Bloquear Caja al Entrar / Desbloquear al Salir
            strSQL = "update pv_cajas set bloqueo = 1" _
                   & " where cod_caja = '" & gCajas.Caja & "' and usuario = '" & gCajas.Usuario & "'"
            Call ConectionExecute(strSQL)
        
         Select Case gCajas.Display
           Case "E01" 'Default
                 Call sbFormsCall("frmPosFacturacion")
           Case "M01" 'Market
                 Call sbFormsCall("frmPosFacturacion")
           Case "VPD" 'Venta de Productos y Articulos
                 Call sbFormsCall("frmPosFacturacion")
           Case "SRV" 'Servicios
                 Call sbFormsCall("frmPosFacturacion_Servicios")
           Case Else
                 Call sbFormsCall("frmPosFacturacion")
         End Select
         
         
        Unload Me
     Else
        MsgBox "Esta caja no tiene apertura existente, debe abrirla primero...", vbExclamation
     End If
     
    
  Else
    MsgBox "La Caja se encuentra Bloqueada...", vbExclamation
  End If 'Bloqueo
  rs.Close
  
 End If 'Select cajas

End If

End Sub
