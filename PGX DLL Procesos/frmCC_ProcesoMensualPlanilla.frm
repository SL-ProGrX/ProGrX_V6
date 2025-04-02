VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCC_ProcesoMensualPlanilla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione la Institución/Deductora a Procesar Rebajos en la Planilla"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10620
      _Version        =   1441793
      _ExtentX        =   18732
      _ExtentY        =   9970
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
   Begin XtremeSuiteControls.CheckBox chkActivas 
      Height          =   252
      Left            =   8040
      TabIndex        =   1
      Top             =   720
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activas?"
      BackColor       =   12582912
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   6255
      _Version        =   1441793
      _ExtentX        =   11033
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Descripción"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Código"
      ForeColor       =   16777215
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmCC_ProcesoMensualPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaInst()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select I.cod_institucion,I.descripcion,I.pr_fecha_corte, isnull(I.Desc_Corta,'') as 'Desc_Corta'" _
       & ", dbo.fxPrm_Deduccion_Aplicada_Fecha(I.pr_Fecha_Corte, I.cod_Institucion) as 'Aplicada'" _
       & ", isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones I inner join Prm_Usuarios U on I.cod_institucion = U.cod_institucion" _
       & " and U.usuario = '" & glogon.Usuario & "'" _
       & " where I.Activa = " & chkActivas.Value

If IsNumeric(txtCodigo.Text) Then
   strSQL = strSQL & " and I.cod_institucion = " & txtCodigo.Text
End If

If Len(Trim(txtDescripcion.Text)) > 0 Then
   strSQL = strSQL & " and I.descripcion like '%" & txtDescripcion.Text & "%'"
End If

If Len(Trim(txtDescripcion.Text)) > 0 Then
   strSQL = strSQL & " OR I.DESC_CORTA like '%" & txtDescripcion.Text & "%'"
End If


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!cod_institucion)
     itmX.SubItems(1) = rs!Desc_Corta
     itmX.SubItems(2) = rs!Descripcion
     
     If rs!Frecuencia_Id = "M" Then
         itmX.SubItems(3) = Year(rs!pr_fecha_Corte) & Format(Month(rs!pr_fecha_Corte), "00")
     End If
     
     If rs!Frecuencia_Id = "Q" Then
        If Day(rs!pr_fecha_Corte) > 15 Then
            itmX.SubItems(3) = Year(rs!pr_fecha_Corte) & Format(Month(rs!pr_fecha_Corte), "00") & ".2"
        Else
            itmX.SubItems(3) = Year(rs!pr_fecha_Corte) & Format(Month(rs!pr_fecha_Corte), "00") & ".1"
        End If
     End If
     
     Select Case rs!Aplicada
        Case 0 'Sin Enviar
            itmX.SubItems(4) = "Pendiente"
            
           itmX.TextBackColor = RGB(249, 212, 155)
        Case 1 'Enviada
            itmX.SubItems(4) = "Enviada"
            itmX.TextBackColor = RGB(239, 249, 155)
        Case 2 'Aplicada
            itmX.SubItems(4) = "Aplicada"
 
     End Select
     
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub chkActivas_Click()
Call sbCargaInst
End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

With lsw.ColumnHeaders
   .Clear
   .Add , , "Id", 600, vbCenter
   .Add , , "Desc.Corta", 1840, vbCenter
   .Add , , "Descripción", 5400
   .Add , , "Proceso", 1140, vbCenter
   .Add , , "Estado", 1140, vbCenter
   
End With

Call sbCargaInst
End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Me.MousePointer = vbHourglass

If lsw.ListItems.Count > 0 Then
  GLOBALES.gInstitucion = Item.Text
  GLOBALES.gNombreInstitucion = Item.SubItems(2)
  GLOBALES.glngFechaCR = Item.SubItems(3)
  
  Call sbFormsCall("frmCC_ProcesoMensual", 0, 0, 0, False, , True)
  
End If

Me.MousePointer = vbDefault

UnLoad Me

End Sub


Private Sub txtCodigo_Change()
Call sbCargaInst
End Sub

Private Sub txtDescripcion_Change()
Call sbCargaInst
End Sub
