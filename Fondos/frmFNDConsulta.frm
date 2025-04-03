VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos a Contratos"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   Icon            =   "frmFNDConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10230
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4212
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   9972
      _Version        =   1572864
      _ExtentX        =   17590
      _ExtentY        =   7429
      _StockProps     =   77
      BackColor       =   -2147483643
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton OptMov 
      Height          =   612
      Index           =   0
      Left            =   6720
      TabIndex        =   14
      Top             =   240
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Aportación"
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
      Appearance      =   16
      Checked         =   -1  'True
      Picture         =   "frmFNDConsulta.frx":030A
   End
   Begin XtremeSuiteControls.PushButton OptMov 
      Height          =   612
      Index           =   1
      Left            =   8400
      TabIndex        =   15
      Top             =   240
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Anulación"
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
      Appearance      =   16
      Picture         =   "frmFNDConsulta.frx":0AF9
   End
   Begin XtremeSuiteControls.FlatEdit txtOperadora 
      Height          =   312
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1092
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtPlan 
      Height          =   312
      Left            =   1200
      TabIndex        =   9
      Top             =   1200
      Width           =   1092
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   312
      Left            =   2280
      TabIndex        =   10
      Top             =   1200
      Width           =   1092
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Left            =   5160
      TabIndex        =   12
      Top             =   1200
      Width           =   4932
      _Version        =   1572864
      _ExtentX        =   8700
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   312
      Left            =   9360
      TabIndex        =   13
      Top             =   5880
      Width           =   732
      _Version        =   1572864
      _ExtentX        =   1291
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "50"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Movimiento a Realizar ....:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   3132
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No Lineas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7200
      TabIndex        =   5
      Top             =   5880
      Width           =   1932
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nombre"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   960
      Width           =   4932
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Identificación"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   1812
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contrato"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Plan"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Operadora"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmFNDConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCajas As Boolean

Private Sub sbActualiza()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem, i As Integer


On Error GoTo vError

lsw.ListItems.Clear

Me.MousePointer = vbHourglass

strSQL = "Select TOP " & txtLineas.Text & " O.Descripcion as Operadora,F.Cod_Operadora,F.Cod_plan,P.Descripcion," _
       & "F.Cod_Contrato,F.Cedula,S.Nombre from Fnd_Contratos F Inner Join Fnd_Operadoras O " _
       & "on F.Cod_operadora=O.Cod_operadora inner join Fnd_planes P " _
       & "on F.Cod_plan=P.Cod_plan inner join Socios S on F.Cedula=S.Cedula " _
       & " Where F.Estado <> 'L' and dbo.fxFndColaboradorVisualiza(F.COD_OPERADORA, F.COD_PLAN, F.cedula, S.ESTADOACTUAL , '" & glogon.Usuario & "') = 1"
 
If Trim(txtOperadora) <> "" Then
   strSQL = strSQL & " And F.Cod_operadora=" & Trim(txtOperadora)
End If

If Trim(txtPlan) <> "" Then strSQL = strSQL & " And F.Cod_Plan like '" & Trim(txtPlan) & "%'"

If Trim(txtContrato) <> "" Then strSQL = strSQL & " And F.Cod_Contrato=" & Trim(txtContrato)

If Trim(txtCedula) <> "" Then
  strSQL = strSQL & " And F.Cedula like '" & Trim(txtCedula) & "%'"
Else
 'Si no aplica la cedula ver por nombre
 If Trim(txtNombre) <> "" Then
    strSQL = strSQL & " And S.nombre like '" & Trim(txtNombre) & "%'"
 End If
End If

Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Operadora)
        itmX.SubItems(1) = Trim(rs!Cod_Plan)
        itmX.SubItems(2) = Trim(rs!Descripcion)
        itmX.SubItems(3) = rs!COD_Contrato
        itmX.SubItems(4) = rs!Cedula
        itmX.SubItems(5) = rs!Nombre
        itmX.Tag = rs!COD_OPERADORA
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Me.imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

vCajas = IIf((fxCajasParametros("03") = "S"), True, False)

With lsw.ColumnHeaders

    .Add , , "Operadora", 1000
    .Add , , "Plan", 1200, vbCenter
    .Add , , "Descripción", 2500
    .Add , , "Contrato", 1200, vbCenter
    .Add , , "Identificación", 1600
    .Add , , "Nombre", 2500
End With


End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Dim sForm As String

If lsw.ListItems.Count = 0 Then Exit Sub

gFondos.Operadora = Item.Tag
gFondos.Plan = Item.SubItems(1)
gFondos.Contrato = Item.SubItems(3)

Select Case True
  Case OptMov.Item(0).Checked
    sForm = "frmCajas_FNDAportaciones"

  Case OptMov.Item(1).Checked
    sForm = "frmFNDAnulaciones"
End Select

Call sbFormsCall(sForm, vbModal, , , False, Me)



End Sub

Private Sub OptMov_Click(Index As Integer)
Dim vCheck As Boolean

vCheck = OptMov.Item(Index).Checked

If Index = 0 Then
   OptMov.Item(0).Checked = IIf(vCheck, False, True)
   OptMov.Item(1).Checked = IIf(OptMov.Item(0).Checked, False, True)
Else
   OptMov.Item(1).Checked = IIf(vCheck, False, True)
   OptMov.Item(0).Checked = IIf(OptMov.Item(1).Checked, False, True)
End If

Call sbActualiza

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   
   txtCedula.SetFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   Call sbActualiza
End If

End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
 Case 48 To 57, 8
 Case vbKeyReturn
  Call sbActualiza
 Case Else
  KeyAscii = 0
End Select

End Sub


Private Sub txtContrato_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
 Case 48 To 57, 8
 Case vbKeyReturn
  Call sbActualiza
 Case Else
  KeyAscii = 0
End Select

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   txtNombre.SetFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   Call sbActualiza
End If

End Sub


Private Sub txtNombre_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
 Case vbKeyReturn
  txtOperadora.SetFocus
  Call sbActualiza
End Select

End Sub


Private Sub txtOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "Descripcion"

   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "select cod_operadora,descripcion from fnd_Operadoras"
   frmBusquedas.Show vbModal

   txtOperadora = gBusquedas.Resultado
   txtOperadora.SetFocus
   gBusquedas.Resultado = ""
   Call sbActualiza
End If

End Sub

Private Sub txtOperadora_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
 Case 48 To 57, 8
 Case vbKeyReturn
  Call sbActualiza
 Case Else
  KeyAscii = 0
End Select

End Sub


Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"

   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"

   gBusquedas.Filtro = " And Cod_operadora=" & Trim(txtOperadora)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_Planes"
   frmBusquedas.Show vbModal

   txtPlan = gBusquedas.Resultado
   txtPlan.SetFocus
   gBusquedas.Resultado = ""
   Call sbActualiza
End If

End Sub

Private Sub txtPlan_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
 Case vbKeyReturn
  Call sbActualiza
End Select

End Sub


