VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_ControlAsgManual 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación Manual : Casos para Gestión de Cobros"
   ClientHeight    =   5904
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10476
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5904
   ScaleWidth      =   10476
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.FlatEdit txtNuevoDesc 
      Height          =   312
      Left            =   4080
      TabIndex        =   13
      Top             =   4200
      Width           =   5532
      _Version        =   1245187
      _ExtentX        =   9758
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   8772
      _Version        =   1245187
      _ExtentX        =   15473
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Filtros para Busqueda: "
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCasosSinAsignar 
         Height          =   252
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   6492
         _Version        =   1245187
         _ExtentX        =   11451
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Mostrar únicamente casos sin asignar !"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkCasosMorosos 
         Height          =   252
         Left            =   2160
         TabIndex        =   6
         Top             =   720
         Width           =   6492
         _Version        =   1245187
         _ExtentX        =   11451
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Mostrar solo casos en Mora !"
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
         Appearance      =   16
         Value           =   1
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAplica 
      Height          =   492
      Left            =   8160
      TabIndex        =   0
      Top             =   5160
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmCO_ControlAsgManual.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkMantener 
      Height          =   252
      Left            =   3960
      TabIndex        =   7
      Top             =   4680
      Width           =   5652
      _Version        =   1245187
      _ExtentX        =   9970
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "&Mantener este expediente asignado a este oficial de cobro"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   1392
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   7332
      _Version        =   1245187
      _ExtentX        =   12933
      _ExtentY        =   2455
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   1812
      _Version        =   1245187
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4080
      TabIndex        =   11
      Top             =   1680
      Width           =   5532
      _Version        =   1245187
      _ExtentX        =   9758
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNuevo 
      Height          =   312
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   4200
      Width           =   1812
      _Version        =   1245187
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtActual 
      Height          =   312
      Left            =   2280
      TabIndex        =   9
      Top             =   3600
      Width           =   7332
      _Version        =   1245187
      _ExtentX        =   12933
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   2
      Left            =   840
      TabIndex        =   3
      Top             =   4200
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Trasladar a.:"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Oficina/Agencia"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1212
      _Version        =   1245187
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Expediente"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Index           =   0
      X1              =   9480
      X2              =   720
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frmCO_ControlAsgManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCedula As String

Private Sub cmdAplica_Click()
Dim strSQL As String

On Error GoTo vError

'Verifica Datos
strSQL = ""

'If txtEstado.Tag = "N" Then strSQL = strSQL & " - Esta persona no se encuentra atrasada en sus cuentas..." & vbCrLf
If Trim(UCase(txtActual.Tag)) = Trim(UCase(txtNuevo.Text)) Then strSQL = strSQL & " - Esta persona ya se encuentra asignada al Ejecutivo?..." & vbCrLf
If txtNuevo.Text = "" Then strSQL = strSQL & " - No se especificó el Ejecutivo de cobro a trasladar..." & vbCrLf
If vCedula = "N" Then strSQL = strSQL & " - No se especificó el expediente..." & vbCrLf

If Len(strSQL) > 0 Then
  MsgBox strSQL, vbExclamation
  Exit Sub
End If

strSQL = "exec spCBRControlAsg '" & vCedula & "','" & txtNuevo & "'," & chkMantener.Value
Call ConectionExecute(strSQL)

txtCedula = ""

MsgBox "Asignación Manual realizada satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnAplicar_Click()

End Sub

Private Sub cmdAplicar_Click()

End Sub

Private Sub Form_Activate()
 vModulo = 4
End Sub

Private Sub Form_Load()
 vModulo = 4
 vCedula = ""
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub

Private Sub txtCedula_Change()
 vCedula = ""
 txtNombre = ""
 txtActual = ""
' txtNuevo = "" 'Conserva el último gestionado.
' txtNuevoDesc = ""
 txtEstado = ""
 txtEstado.Tag = "N"
 chkMantener.Value = vbChecked
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select Soc.Cedula,Soc.Nombre,Reg.Id_Solicitud as 'Operación', Reg.Codigo as 'Línea', Cat.Descripcion as 'Línea - Desc.'" _
                        & " from socios Soc inner join Reg_Creditos Reg on Soc.Cedula = Reg.Cedula and Reg.Estado = 'A'" _
                        & " inner join Catalogo Cat on Reg.Codigo = Cat.Codigo and Cat.LINEA_INTERNA = 1" _
                        & " left join Vista_Morosidad Vm on Reg.id_Solicitud = Vm.id_Solicitud"
    gBusquedas.Columna = "Soc.Cedula"
    gBusquedas.Orden = "Soc.cedula"
    gBusquedas.Filtro = ""
    If chkCasosSinAsignar.Value = vbChecked Then
        gBusquedas.Filtro = " and Soc.Cedula not in(select cedula from CBR_ASIGNACION)"
    End If
    
    If chkCasosMorosos.Value = vbChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and isnull(Vm.Id_Solicitud,0) > 1"
    End If
    
    
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado2)
End If

End Sub



Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtNombre = fxNombre(txtCedula)
If txtNombre <> "" Then vCedula = txtCedula

Call sbCBRControlEstado(txtCedula, txtEstado)

strSQL = "select * from cbr_asignacion where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   txtActual.Text = "** Este Expediente no ha sido asignado a ningún oficial **"
   txtActual.Tag = ""
Else
   txtActual.Text = "Oficial : " & rs!Usuario & " / Fecha : " & Format(rs!fecha_asignacion, "dd/mm/yyyy") _
         & " / Mantener : " & IIf((rs!mantener = 1), "SI", "NO")
   txtActual.Tag = Trim(rs!Usuario)

End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEstado.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select Soc.Cedula,Soc.Nombre,Reg.Id_Solicitud as 'Operación', Reg.Codigo as 'Línea', Cat.Descripcion as 'Línea - Desc.'" _
                        & " from socios Soc inner join Reg_Creditos Reg on Soc.Cedula = Reg.Cedula and Reg.Estado = 'A'" _
                        & " inner join Catalogo Cat on Reg.Codigo = Cat.Codigo and Cat.LINEA_INTERNA = 1" _
                        & " left join Vista_Morosidad Vm on Reg.id_Solicitud = Vm.id_Solicitud"
    gBusquedas.Columna = "Soc.Nombre"
    gBusquedas.Orden = "Soc.Nombre"
    gBusquedas.Filtro = ""
    If chkCasosSinAsignar.Value = vbChecked Then
        gBusquedas.Filtro = " and Soc.Cedula not in(select cedula from CBR_ASIGNACION)"
    End If
    
    If chkCasosMorosos.Value = vbChecked Then
        gBusquedas.Filtro = gBusquedas.Filtro & " and isnull(Vm.Id_Solicitud,0) > 1"
    End If
    
    
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = Trim(gBusquedas.Resultado2)
    txtCedula.SetFocus
End If


End Sub

Private Sub txtNuevo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNuevoDesc.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "usuario"
    gBusquedas.Orden = "usuario"
    gBusquedas.Filtro = " and estado = 1"
    frmBusquedas.Show vbModal
    txtNuevo = Trim(gBusquedas.Resultado)
    txtNuevoDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtNuevoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkMantener.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select usuario,nombre from cbr_usuarios"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Filtro = " and estado = 1"
    frmBusquedas.Show vbModal
    txtNuevo = Trim(gBusquedas.Resultado)
    txtNuevoDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

