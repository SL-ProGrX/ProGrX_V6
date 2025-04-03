VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAF_CD_PreCalculo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pre-Cálculo "
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswActividades 
      Height          =   2295
      Left            =   2040
      TabIndex        =   12
      Top             =   3000
      Width           =   8295
      _Version        =   1572864
      _ExtentX        =   14631
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnActividades 
      Height          =   375
      Left            =   9000
      TabIndex        =   22
      Top             =   2620
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actividades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtComiteId 
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   582
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtComiteDesc 
      Height          =   330
      Left            =   3120
      TabIndex        =   1
      Top             =   1200
      Width           =   7215
      _Version        =   1572864
      _ExtentX        =   12726
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtAsociados 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAjusteAsoc 
      Height          =   330
      Left            =   5640
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   582
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAsocTotalAjustado 
      Height          =   330
      Left            =   9240
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   582
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboActividadTipo 
      Height          =   330
      Left            =   2040
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
      _Version        =   1572864
      _ExtentX        =   6376
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtFechaLiq 
      Height          =   330
      Left            =   5040
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
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
      BackColor       =   16777215
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFechaRegistro 
      Height          =   330
      Left            =   8640
      TabIndex        =   17
      Top             =   6000
      Visible         =   0   'False
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
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
      BackColor       =   16777215
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtMontoPagar 
      Height          =   330
      Left            =   2040
      TabIndex        =   18
      Top             =   6000
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   21
      Top             =   6000
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Liq."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a Pagar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   6000
      Width           =   1560
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo de Actividad"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   615
      Index           =   11
      Left            =   240
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Selecciones las actividades"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   5400
      Width           =   10575
      _Version        =   1572864
      _ExtentX        =   18653
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Resumen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   10575
      _Version        =   1572864
      _ExtentX        =   18653
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Actividades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pre-Cálculo de Cuentas a Comités"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   4155
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Comité"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asociados por"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asociados de Ajuste"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Total de Asociados"
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
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmAF_CD_PreCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Dim itmX As ListViewItem
Dim vOperacion As Long


Dim Miembro As String, vCodigo As String
Dim i As Integer, x As Integer

Dim vFecha As Date

Function fxNomComite(vUnidad As String)
   
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select U.descripcion from uprogramatica U right join afi_cd_comites_unidades A " _
            & "on U.codigo = A.cod_comite" _
            & " where A.cod_comite = '" & vUnidad & "'"
            Call OpenRecordSet(rs, strSQL)
   If rs.EOF Then
      fxNomComite = "No existe unidad definida en Comites y Delegados"
   Else
      fxNomComite = rs!Descripcion
   End If
rs.Close
End Function


Private Sub sbCalMiembros()

On Error GoTo vError

   
strSQL = "select count(*) as 'Cantidad' from socios" _
       & " where EstadoActual = 'S' and cod_departamento in(select Codigo_UP from Afi_CD_Comites_Unidades where cod_comite = '" & txtComiteId.Text & "')"
Call OpenRecordSet(rs, strSQL)

txtAsociados.Text = rs!Cantidad
txtAsocTotalAjustado.Text = rs!Cantidad - CCur(txtAjusteAsoc.Text)

rs.Close


Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  
End Sub

Sub sbLlamaComite(vComite As String)

Dim strSQL As String
Dim rs As New ADODB.Recordset
              
     Call cboActividadTipo_Click
                
     strSQL = "select N.cedula, S.nombre, N.cod_comite, C.descripcion" _
            & " from afi_cd_comites C left join afi_cd_nombramientos N on C.cod_comite = N.cod_comite" _
            & " inner join socios S on S.cedula = N.cedula " _
            & " where N.cod_comite = '" & vComite & "' and N.APL_DESEMBOLSOS = 1"
     Call OpenRecordSet(rs, strSQL)
     
     If Not rs.EOF Then
        txtComiteDesc.Text = Trim(rs!Descripcion)
     Else
        MsgBox "No se cuenta con miembro asigando el desembolso!!"
     End If
     rs.Close
     
End Sub


Private Sub sbLimpiar()
     
     vOperacion = 0
     
     txtComiteId.Text = ""
     txtComiteDesc.Text = ""
     
     txtAsociados.Text = "0"
     txtAjusteAsoc.Text = 0
     txtAsocTotalAjustado.Text = "0"
     txtMontoPagar.Text = 0
     
     txtFechaLiq.Text = ""
     txtFechaRegistro.Text = ""
     
     vPaso = False
'     txtComiteId.SetFocus

End Sub


Private Sub sbTiposActividadActiva()

txtFechaLiq.Text = ""
txtMontoPagar.Text = 0

End Sub

Private Sub btnActividades_Click()
 Call sbFormsCall("frmAF_CD_Actividades")
End Sub

Private Sub cboActividadTipo_Click()
If vPaso Then Exit Sub

Dim curMonto As Currency

On Error GoTo vError

curMonto = 0


strSQL = "exec spAFI_CD_Actividades_List '" & cboActividadTipo.ItemData(cboActividadTipo.ListIndex) & "', " & txtAsocTotalAjustado.Text _
       & ", " & vOperacion & ", '" & txtComiteId.Text & "'"
Call OpenRecordSet(rs, strSQL)


vPaso = True

With lswActividades.ListItems
    .Clear
    
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Cod_Actividad)
            itmX.SubItems(1) = RTrim(rs!Descripcion)
            itmX.SubItems(2) = Format(rs!Monto, "Standard")
            itmX.SubItems(3) = rs!Tipo
            
            If rs!Asignado = 1 Then
                itmX.Checked = True
                curMonto = curMonto + rs!Monto
            End If
        rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

txtMontoPagar.Text = Format(curMonto, "Standard")

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
   
 vModulo = 40
   
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
   
 vFecha = fxFechaServidor
 
 vPaso = True

 strSQL = "select CodTipoActividad as 'Idx', NombreTipoActividad as 'ItmX' from AFI_CD_TIPO_ACTIVIDAD where Activo = 1"
 Call sbCbo_Llena_New(cboActividadTipo, strSQL, False, True)
  
 With lswActividades.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Actividad", 3200
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Tipo", 2200
 End With
 lswActividades.Checkboxes = True
 
 
vPaso = False
 
 txtFechaRegistro.Text = Format(vFecha, "dd/mm/yyyy")
 
 Call sbTiposActividadActiva
 
 Call sbLimpiar

End Sub

Private Sub lswActividades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub


If Item.Checked Then
    txtMontoPagar.Text = Format(CCur(txtMontoPagar.Text) + CCur(Item.SubItems(2)), "Standard")
Else
    txtMontoPagar.Text = Format(CCur(txtMontoPagar.Text) - CCur(Item.SubItems(2)), "Standard")
End If


End Sub


Private Sub txtAjusteAsoc_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtAsocTotalAjustado.SetFocus
End If

Exit Sub

vError:

End Sub



Private Sub txtAjusteAsoc_LostFocus()

On Error GoTo vError

txtAsocTotalAjustado.Text = CCur(txtAsociados.Text) - CCur(txtAjusteAsoc.Text)
Call cboActividadTipo_Click

Exit Sub

vError:

End Sub

Private Sub txtComiteId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_COMITE, DESCRIPCION from AFI_CD_COMITES"
       gBusquedas.Filtro = " AND ACTIVO = 1"
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtComiteId.Text = gBusquedas.Resultado
       Call txtComiteId_KeyPress(vbKeyReturn)
End If
End Sub

Private Sub txtComiteId_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
         txtComiteDesc.SetFocus
      Case Else
       KeyAscii = 0
  End Select
End Sub


Private Sub txtComiteDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select distinct cod_comite,U.descripcion from afi_cd_comites_unidades A " _
                             & "left join uprogramatica U on A.cod_comite = U.codigo "
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtComiteId.Text = gBusquedas.Resultado
       Call txtComiteId_KeyPress(vbKeyReturn)
 
End If

End Sub

Private Sub txtComiteId_LostFocus()
        Call sbLlamaComite(txtComiteId.Text)
        Call sbCalMiembros
        Call cboActividadTipo_Click
End Sub
