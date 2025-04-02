VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmVivDetalleGarantia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de acreedores de hipotecas"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lvwDetalle 
      Height          =   2295
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   7575
      _Version        =   1310723
      _ExtentX        =   13361
      _ExtentY        =   4048
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin ComCtl3.CoolBar Clb_Barra_Tareas 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   688
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   9975
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   3825
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtObservaciones 
      Height          =   1275
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   7575
      _Version        =   1310723
      _ExtentX        =   13361
      _ExtentY        =   2249
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPropietario 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   7575
      _Version        =   1310723
      _ExtentX        =   13361
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   2175
      _Version        =   1310723
      _ExtentX        =   3836
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
      Text            =   "0.00"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboGradoHipoteca 
      Height          =   330
      Left            =   6600
      TabIndex        =   6
      Top             =   960
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
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
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   11
      Top             =   960
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Grado de la Hipoteca"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   7575
      _Version        =   1310723
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Registradas: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Observaciones"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Monto"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Acreedor"
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmVivDetalleGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim vEditar As Boolean

Public m_DescGradoH As String
Private m_cambioDatos  As Boolean

Public m_IdGarantia As Long
Public m_IdLinea As Integer

Private Sub sbLimpiarCampos()
txtPropietario.Text = Empty
txtMonto.Text = Format(0, "Standard")
txtObservaciones.Text = Empty
m_cambioDatos = False

End Sub

Public Sub sbListaDetalle(ByVal pIdGarantia As Long, ByVal pLinea As Integer)

Dim vItem As ListViewItem

Dim vKey As String

On Error GoTo vError

lvwDetalle.ListItems.Clear
lvwDetalle.ColumnHeaders.Clear
lvwDetalle.ColumnHeaders.Add , , "Idgarantia", 0
lvwDetalle.ColumnHeaders.Add , , "linea", 0
lvwDetalle.ColumnHeaders.Add , , "Propietario", 4000
lvwDetalle.ColumnHeaders.Add , , "Monto", 2000, 1
lvwDetalle.ColumnHeaders.Add , , "Grado Hipoteca", 2000
lvwDetalle.ColumnHeaders.Add , , "Observaciones", 3000

strSQL = "SELECT IdGarantia, Linea, Propietario, Monto, " & _
        " Case GradoHipoteca " & _
        " when 'P' then 'Primer Grado' " & _
        " when 'S' then 'Segundo Grado' " & _
        " when 'T' then 'Tercer Grado'" & _
        " end As DescGradoHiporteca, observaciones " & _
        " From ViviendaGarantiaDetalle "
                

If pLinea = -1 Then 'Consulta todas la lineas detalle registras al numero de garantia
    strSQL = strSQL & " WHERE (IdGarantia = " & pIdGarantia & ")"
Else 'Consulta  una linea detalle registra al numero de garantia
    strSQL = strSQL & " WHERE (IdGarantia = " & pIdGarantia & ") and (linea = " & pLinea & ")"""
End If

Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then

Do While Not rs.EOF
    vKey = "(VV)" & Trim(rs!IdGarantia) _
           & "(Ig)" & Trim(rs!Linea) & "(Ln)"
           
    Set vItem = lvwDetalle.ListItems.Add(, vKey, rs!IdGarantia)
        vItem.SubItems(1) = rs!Linea
        vItem.SubItems(2) = Trim(rs!Propietario)
        vItem.SubItems(3) = Format(rs!Monto, "Standard")
        vItem.SubItems(4) = IIf(IsNull(rs!DescGradoHiporteca), "", Trim(rs!DescGradoHiporteca))
        vItem.SubItems(5) = IIf(IsNull(rs!Observaciones), "", Trim(rs!Observaciones))
        
       rs.MoveNext
Loop
rs.Close
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vEditar = False

    m_DescGradoH = GLOBALES.gTag
    m_IdGarantia = GLOBALES.gTag2
    m_IdLinea = GLOBALES.gTag3


Call sbToolBarIconos(tlbPrincipal, False)
Call sbToolBar(tlbPrincipal, "nuevo")


    Call sbCargaGradoHiporteca

Call sbListaDetalle(m_IdGarantia, -1)
 
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Me.ActiveControl.Name = "txtObservaciones" Then Exit Sub

If (KeyCode = vbKeyReturn) Then
    Call gsbPulsarTecla(vbKeyTab)
End If

End Sub


Private Sub sbCargaGradoHiporteca()

    Dim i As Integer
    cboGradoHipoteca.Clear
    
    Select Case m_DescGradoH
        Case "Segundo Grado"
                cboGradoHipoteca.AddItem "Primer Grado"
                
                cboGradoHipoteca.Text = "Primer Grado"
        Case "Tercer Grado"
                cboGradoHipoteca.AddItem "Primer Grado"
                cboGradoHipoteca.AddItem "Segundo Grado"
                
                cboGradoHipoteca.Text = "Primer Grado"
    End Select


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo vError

If Me.ActiveControl.Name = "txtMonto" Then
    KeyAscii = fxValidaEnterosYDecimal(Trim$(txtMonto.Text), KeyAscii)
End If

Exit Sub

vError:
    MsgBox "Ocurrió un error validar la información de los formatos. " & "-" & Err.Description, vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)


m_DescGradoH = Empty
m_IdGarantia = 0
m_IdLinea = 0

End Sub



Private Sub lvwDetalle_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
m_IdGarantia = fxDeCodePK(Item.Key, 5, "(Ig)")
m_IdLinea = fxDeCodePK(Item.Key, gPosIni, "(Ln)")

txtPropietario.Text = Item.SubItems(2)
txtMonto.Text = Format(Item.SubItems(3), "Standard")

cboGradoHipoteca.Text = Trim(Item.SubItems(4))

m_DescGradoH = Item.SubItems(4)
txtObservaciones.Text = Trim(Item.SubItems(5))

Call sbToolBar(Me.tlbPrincipal, "Activo")
End Sub



Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

    Select Case LCase(Button.Key)
    
        Case "nuevo"
            vEditar = False
            Call sbToolBar(Me.tlbPrincipal, "edicion")
            Call sbLimpiarCampos
            
        Case "editar"
            vEditar = True
            Call sbToolBar(Me.tlbPrincipal, "edicion")
            
        Case "borrar"
            Call sbBorrar
            
        Case "guardar"
            Call sbGuardar
            
        Case "deshacer"
            vEditar = False
            Call sbToolBar(Me.tlbPrincipal, "nuevo")
             Call sbLimpiarCampos
    End Select
    

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
    
End Sub
Private Function fxEstadoOperacion(ByVal IdGarantia As Long) As String

On Error GoTo vError

fxEstadoOperacion = ""
                
strSQL = "SELECT R.ESTADOSOL" & _
        " FROM   ViviendaGarantia as G INNER JOIN ViviendaGarantiaDetalle as Detalle ON G.IdGarantia = Detalle.IdGarantia" & _
        " INNER JOIN REG_CREDITOS AS R ON G.NumeroOperacion = R.ID_SOLICITUD" & _
        " and  G.Idgarantia = " & IdGarantia
Call OpenRecordSet(rs, strSQL)
                       
If Not glogon.error Then
    fxEstadoOperacion = rs!ESTADOSOL
    rs.Close
End If

Exit Function

vError:
MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Function

Private Sub sbBorrar()
On Error GoTo vError

 If vEditar = False Then
     If fxEstadoOperacion(m_IdGarantia) = "F" Then
        Me.MousePointer = vbDefault
        MsgBox "No es posible realizar movimientos para un número de operación en estado FORMALIZADA.", vbExclamation
        Exit Sub
    End If
    
    strSQL = "delete ViviendaGarantiaDetalle where IdGarantia = " & m_IdGarantia & "  and Linea = " & m_IdLinea
    Call ConectionExecute(strSQL)
    If Not glogon.error Then
        MsgBox "La información seleccionada fue borrada correctamente", vbInformation
        Call sbLimpiarCampos
        Call sbListaDetalle(m_IdGarantia, -1)
        Call sbToolBar(Me.tlbPrincipal, "nuevo")
    End If
End If
 

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbGuardar()
If m_cambioDatos = False Then Exit Sub

If Not vEditar Then 'Nuevo
    Call sbAgregar
Else 'Nuevo
    Call sbModificar
End If

End Sub

Private Sub sbAgregar()

On Error GoTo vError

Me.MousePointer = vbHourglass

If fxValidaDatos = False Then Exit Sub

'pIdGarantia As Long, ByVal pLinea As Integer, ByVal pPropietario As String, _
'                                          ByVal pMonto As String, ByVal pGradoHipoteca As String, _
'                                          ByVal pObservaciones As String, ByVal pRegistroUsuario As String) As Boolean

strSQL = "exec spCRDVivGarantiaDetalle_A " & gParametros(0) & ",-1,'" & gParametros(1) _
       & "'," & CCur(gParametros(2)) & ",'" & gParametros(3) & "','" & gParametros(4) & "','" & gParametros(5) & "'"
Call ConectionExecute(strSQL)
If Not glogon.error Then
    m_cambioDatos = False
    
    Call Bitacora("REGISTRA", "Garantias vivienda hipoteca: " & gParametros(0) & " monto: " & gParametros(2))
    
    MsgBox "Información fue registrada corretamente.", vbInformation
    
    Call sbListaDetalle(m_IdGarantia, -1)
    Call sbLimpiarCampos
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
    
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
    
End Sub

Private Sub sbModificar()

On Error GoTo error

Me.MousePointer = vbHourglass

If fxValidaDatos = False Then Exit Sub

strSQL = "exec spCRDVivGarantiaDetalle_A " & gParametros(0) & "," & m_IdLinea & ",'" & gParametros(1) _
       & "'," & CCur(gParametros(2)) & ",'" & gParametros(3) & "','" & gParametros(4) & "','" & gParametros(5) & "'"
Call ConectionExecute(strSQL)
If Not glogon.error Then
    m_cambioDatos = False
    
    Call Bitacora("MODIFICA", "Garantias vivienda hipoteca: " & gParametros(0) & " monto: " & gParametros(2))
    
    MsgBox "Información fue actualizada correctamente.", vbInformation
    Call sbListaDetalle(m_IdGarantia, -1)
    Call sbLimpiarCampos
    Call sbToolBar(Me.tlbPrincipal, "nuevo")
    
End If
 
Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    Call ObjMensajes.deError("Ocurrió un error en visual basic al modificar la información ingresada. Error " & Err.Description)
    
End Sub

Private Function fxValidaDatos() As Boolean

On Error GoTo vError

fxValidaDatos = False

ReDim gParametros(0 To 5)

If fxEstadoOperacion(m_IdGarantia) = "F" Then
    Me.MousePointer = vbDefault
    MsgBox ("No es posible realizar movimientos para un número de operación en estado FORMALIZADA."), vbExclamation
    Exit Function
End If

If (Len(Trim(txtPropietario.Text)) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un nombre para el propietario."), vbExclamation
    txtPropietario.SetFocus
    Exit Function
End If

If (Val(txtMonto.Text)) = 0 Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar un monto válido."), vbExclamation
    txtMonto.SetFocus
    Exit Function
End If

Select Case Trim(cboGradoHipoteca.Text)
    Case "Primer Grado"
        gParametros(3) = "P"
    Case "Segundo Grado"
        gParametros(3) = "S"
    Case "Tercer Grado"
        gParametros(3) = "T"
End Select

gParametros(0) = m_IdGarantia
gParametros(1) = Trim(txtPropietario.Text)
gParametros(2) = CCur(txtMonto.Text)
gParametros(4) = IIf((Len(Trim(txtObservaciones.Text)) = 0), "", Trim(txtObservaciones.Text))
gParametros(5) = glogon.Usuario

fxValidaDatos = True
Exit Function

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Function


Private Sub txtMonto_GotFocus()
If Val(txtMonto.Text) = 0 Then Exit Sub
txtMonto.Text = CCur(txtMonto.Text)
End Sub

Private Sub txtMonto_Change()
m_cambioDatos = True
End Sub

Private Sub txtMonto_LostFocus()
If Len(txtMonto.Text) = 0 Then Exit Sub
txtMonto.Text = Format(txtMonto.Text, "Standard")
End Sub

Private Sub TxtObservaciones_Change()
m_cambioDatos = True
End Sub

Private Sub txtPropietario_Change()
m_cambioDatos = True
End Sub

Private Sub txtValorTerreno_Change()
m_cambioDatos = True
End Sub

Private Sub cboGradoHipoteca_Click()
m_cambioDatos = True
End Sub

