VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmVivHonorariosDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle de honorarios"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAceptar 
      Height          =   492
      Left            =   6000
      TabIndex        =   6
      Top             =   5640
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
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
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2532
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   8412
      _Version        =   524288
      _ExtentX        =   14838
      _ExtentY        =   4466
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   486
      MaxRows         =   498
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmVivHonorariosDetalle.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdCancelar 
      Height          =   492
      Left            =   7320
      TabIndex        =   7
      Top             =   5640
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
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
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   672
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1185
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   16.5
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1800
      TabIndex        =   9
      Top             =   960
      Width           =   1812
      _Version        =   1441793
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedulaContacto 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3600
      TabIndex        =   11
      Top             =   960
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
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
   Begin XtremeSuiteControls.FlatEdit txtNombreContacto 
      Height          =   312
      Left            =   3600
      TabIndex        =   12
      Top             =   1320
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deudor: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Profesional:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " Total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   15
      Left            =   5160
      TabIndex        =   3
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblTotalMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVivHonorariosDetalle.frx":070C
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   588
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   8508
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
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
      Index           =   6
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12372
   End
End
Attribute VB_Name = "frmVivHonorariosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_IdDesembolso As Long
Public m_IdContacto As Long
Public m_IdGarantia As Long
Public m_Profesional As String
Private m_Guardar As Boolean

Private Function sbAgregarDetalle() As Boolean
Dim vCodigo As String, strSQL As String
Dim i As Integer
Dim vMonto As Double

'Inicia Proceso
Me.MousePointer = vbHourglass

On Error GoTo vError
sbAgregarDetalle = True

If ObjAgregar.fxViviendaAsingarGarantia(m_IdGarantia, m_IdContacto, m_Profesional, glogon.Usuario, "1900/01/01") Then
    frmVivControlAsignacionGarantia.m_AsignaGarantia = True
    
    Call Bitacora("APLICA", "Asignación Garantia Vivienda: " & m_IdGarantia & " Contacto: " & m_IdContacto)
    
End If

'Inicia registro de detalle
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.Col = 2
    If CCur(vGrid.Text) > 0 Then
        vMonto = CCur(vGrid.Text)
        vGrid.Col = 3
        vCodigo = vGrid.Text
       
        If Not ObjAgregar.fxHonorariosDetalle(m_IdContacto, m_IdGarantia, m_Profesional, vCodigo, vMonto, glogon.Usuario) Then
            m_Guardar = False
            Exit For
        End If
        
        m_Guardar = True
    End If
Next i
' Eliminado al cambiar el llamado de la pantalla del grid registro, para el grid de profesionales
If m_Guardar Then
'    Call ObjActualizar.fxCtlAsignacionGarantia(m_IdGarantia, m_IdContacto, glogon.usuario, "S", "A", "I")
    strSQL = "exec spCRDViviendaDesembolsoPendiente " & m_IdContacto & ",'A'," & m_IdGarantia & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    MsgBox "Información fue registrada corretamente.", vbInformation
End If

Me.MousePointer = vbDefault
Exit Function
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function

Private Sub cmdAceptar_Click()

If m_Guardar Then
    MsgBox ("Los datos ya fueron almacenados, no es posible realizar modificaciones.")
    Exit Sub
End If
If (MsgBox("¿ Confirma que desea guardar la información suministrada para el detalle de honorarios.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub

If Val(lblTotalMonto.Caption) = 0 Then
    If (MsgBox("¿ Los la suma de los montos digitados no son mayores a cero, esta seguro que desea continuar con el proceso.?", vbQuestion + vbYesNo) = vbNo) Then Exit Sub
End If

Call sbAgregarDetalle
UnLoad Me

End Sub

Private Sub cmdCancelar_Click()
UnLoad Me
End Sub

Private Sub Form_Load()
            
vModulo = 3
  
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
  
m_IdContacto = GLOBALES.gTag
m_IdGarantia = GLOBALES.gTag2
m_Profesional = GLOBALES.gTag3
            
vGrid.AppearanceStyle = fxGridStyle

Call sbCargaInfoOperacion(gOperacion)
Call sbHonorariosDetalle

m_Guardar = False
    
    
End Sub

Public Sub sbCargaInfoOperacion(ByVal pOperacion As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

  glogon.strSQL = "select R.id_solicitud,R.cedula,S.nombre from reg_creditos R inner join Socios S" _
          & " on R.cedula = S.cedula where R.id_solicitud = " & pOperacion
   
If execSql(glogon.strSQL) Then
  txtOperacion.Text = glogon.Recordset!Id_solicitud
  txtCedula.Text = RTrim(glogon.Recordset!cedula)
  txtNombre.Text = glogon.Recordset!Nombre
End If


    glogon.strSQL = "select isnull(identificacion,'') as identificacion, isnull(nombre,'') as nombre " _
              & " from viviendacontactos where idcontacto = " & m_IdContacto
       
    If execSql(glogon.strSQL) Then
      txtCedulaContacto.Text = RTrim(glogon.Recordset!Identificacion)
      txtNombreContacto.Text = glogon.Recordset!Nombre
    End If


 
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
Private Sub sbSumaLineas()

On Error GoTo vError

Dim i As Integer
Dim vIntActuales As Double
Dim vDisponible As Double

lblTotalMonto.Caption = Format(0, "Standard")

For i = 1 To vGrid.MaxRows
    vGrid.Col = 2
    vGrid.Row = i
    If CCur(vGrid.Text) > 0 Then
        lblTotalMonto.Caption = CCur(lblTotalMonto.Caption) + CCur(vGrid.Text)
    End If
Next i
  
lblTotalMonto.Caption = Format(lblTotalMonto.Caption, "Standard")
  
Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub
Private Sub sbHonorariosDetalle()
vGrid.ColWidth(3) = 0
If ObjConsultar.fxHonorariosDetalle_TT Then
    Call sbCargavGridLocal(vGrid, 3)
    lblTotalMonto.Caption = Format(0, "Standard")
End If

End Sub
Private Sub sbCargavGridLocal(vGrid As Object, vGridMaxCol As Integer)
Dim rs As New ADODB.Recordset, i As Integer
Set rs = glogon.Recordset
vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1).Value))
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
vGrid.MaxRows = vGrid.MaxRows - 1

vGrid.SetActiveCell 2, vGrid.ActiveRow

rs.Close

End Sub


Private Sub vGrid_EditChange(ByVal Col As Long, ByVal Row As Long)
Call sbSumaLineas
End Sub
