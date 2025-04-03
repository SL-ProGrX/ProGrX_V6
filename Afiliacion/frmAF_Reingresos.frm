VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Reingresos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8235
   Icon            =   "frmAF_Reingresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.FlatEdit txtCodPromotor 
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2160
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtNombrePromotor 
      Height          =   372
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   2160
      Width           =   6012
      _Version        =   1441793
      _ExtentX        =   10604
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtBoleta 
      Height          =   372
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Visible         =   0   'False
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   656
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   0
      Left            =   5400
      TabIndex        =   6
      Top             =   2880
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Reingresos.frx":000C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   1
      Left            =   6720
      TabIndex        =   7
      Top             =   2880
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_Reingresos.frx":0733
      ImageAlignment  =   4
   End
   Begin XtremeShortcutBar.ShortcutCaption lblNombre 
      Height          =   372
      Left            =   2040
      TabIndex        =   9
      Top             =   960
      Width           =   6252
      _Version        =   1441793
      _ExtentX        =   11028
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption lblCedula 
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblMovimiento 
      Height          =   972
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8292
      _Version        =   1441793
      _ExtentX        =   14626
      _ExtentY        =   1714
      _StockProps     =   14
      Caption         =   "Registro de Afiliación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.45
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Boleta N°"
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
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Promotor"
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
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   852
   End
End
Attribute VB_Name = "frmAF_Reingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vTipoMovimento As String

Private Sub btnTool_Click(Index As Integer)


Select Case Index
    Case 0 'Aplicar
      If fxValida Then
        If vTipoMovimento = "R" Then
          Call sbReingresa
        Else
          Call sbActivar
        End If
      End If

    Case 1 'Cancelar
      Unload Me
End Select

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()
vModulo = 1

vTipoMovimento = Mid(GLOBALES.gTag2, Len(GLOBALES.gTag2), 1)
'If vTipoMovimento = "R" Then
'  lblMovimiento.Caption = "ReIngresar"
'Else
' lblMovimiento.Caption = "Activar"
'End If

lblCedula.Caption = Mid(GLOBALES.gTag2, 1, Len(GLOBALES.gTag2) - 1)
GLOBALES.gTag2 = lblCedula.Caption
lblNombre.Caption = GLOBALES.gTag3

txtBoleta.Text = "0"

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCodPromotor.SetFocus
End Sub

Private Sub txtCodPromotor_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   txtCodPromotor = ""
   txtNombrePromotor = ""
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select ID_PROMOTOR ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtCodPromotor = Trim(gBusquedas.Resultado)
   txtNombrePromotor = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtNombrePromotor_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   txtNombrePromotor.Text = ""
   txtCodPromotor.Text = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   
   frmBusquedas.Show vbModal
   txtCodPromotor.Text = Trim(gBusquedas.Resultado)
   txtNombrePromotor.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub sbReingresa()
Dim i As Integer

i = MsgBox("Está seguro que desea Re Ingresar a esta Persona?", vbYesNo)
        
If i = vbYes Then
    Call sbReIngreso(GLOBALES.gTag2, Val(GLOBALES.gTag), txtCodPromotor.Text, txtBoleta.Text, fxFechaServidor)
    
    Call Bitacora("Registra", "Re-Ingreso de Persona - Cédula:" & GLOBALES.gTag2)
    
    If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("03", "Re-Ingreso de Persona - Cedula: " & Trim(GLOBALES.gTag2), Trim(GLOBALES.gTag2))
    End If
     Call sbSIFRegistraTags(lblCedula.Caption, "S02", "Afiliación", fxgAFIIngresoConsecutivo(lblCedula.Caption, fxFechaServidor), "AFI")
     
    MsgBox "Persona Reingresada Satisfactoriamente...", vbInformation
    Unload Me
End If

End Sub

Private Sub sbInsertAhorro()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAFI_PERSONA_PATRIMONIO_Vincula '" & Trim(GLOBALES.gTag2) & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String, strSQL As String, rs As New ADODB.Recordset

vMensaje = ""

If Trim(txtBoleta.Text) = "" Then
  vMensaje = vMensaje & vbCrLf & " - Número de Boleta Invalido..."
End If
If Trim(txtCodPromotor.Text) = "" Then
  vMensaje = vMensaje & vbCrLf & " - Promotor invalido Boleta Invalido..."
Else
strSQL = "select isnull(estado,0)as estado from promotores where id_promotor = " & txtCodPromotor & ""
Call OpenRecordSet(rs, strSQL)
If rs!Estado = 0 Then
    vMensaje = vMensaje & vbCrLf & " - El promotor indicado se encuenta inactivo o no existe..."
End If
End If
If Len(vMensaje) = 0 Then
   fxValida = True
Else
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function

Private Sub sbActivar()
Dim i As Integer, strSQL As String, vFecha As Date

i = MsgBox("Esta seguro que desea Activar a Esta Persona", vbYesNo)

If i = vbYes Then


  vFecha = fxFechaServidor
  
  strSQL = "update socios set estadoactual = 'S',FechaIngreso = '" & Format(vFecha, "yyyy/mm/dd") & "'" _
         & ",priDeduc = " & fxgPrimerDeduccionIng(Val(GLOBALES.gTag)) _
         & ",reg_user = '" & glogon.Usuario & "',reg_fecha = dbo.MyGetdate(), Fecha_Comision = Null" _
         & ",id_promotor = " & txtCodPromotor & ", cod_oficina = '" & GLOBALES.gOficinaTitular & "' where cedula = '" & Trim(GLOBALES.gTag2) & "'"
  Call ConectionExecute(strSQL)


 'Procesa Historico de Ingreso
  strSQL = "Insert afi_ingresos(Cedula, fecha_ingreso, id_promotor, Boleta, Usuario, Fecha,cod_oficina)" _
         & " values('" & Trim(GLOBALES.gTag2) & "',dbo.MyGetdate()," _
         & txtCodPromotor.Text & ",'" & txtBoleta & "','" & glogon.Usuario _
         & "',dbo.MyGetdate(),'" & GLOBALES.gOficinaTitular & "')"
  Call ConectionExecute(strSQL)


  Call sbInsertAhorro

  Call Bitacora("Registra", "Activacion de Persona - Cedula:" & GLOBALES.gTag2)

  If vParametros.BitacoraEspecial Then
     Call sbgAFIBitacora("04", "Activacion de Persona - Cedula " & Trim(GLOBALES.gTag2), Trim(GLOBALES.gTag2))
  End If
  Call sbSIFRegistraTags(lblCedula.Caption, "S07", "Afiliación", fxgAFIIngresoConsecutivo(lblCedula.Caption, fxFechaServidor), "AFI")
 

  MsgBox "Persona se Activo Satisfactoriamente...", vbInformation
  Unload Me
End If
End Sub
