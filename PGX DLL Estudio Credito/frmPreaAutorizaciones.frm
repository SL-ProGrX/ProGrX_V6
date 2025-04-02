VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmPreaAutorizaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorizaciones por Comité"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8430
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lswAutorizadores 
      Height          =   3372
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8172
      _Version        =   1441792
      _ExtentX        =   14414
      _ExtentY        =   5948
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
      Appearance      =   16
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Indique los Autorizadores"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   432
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   5688
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Private mComite As Integer

Private Sub Form_Load()
    
Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
    
With lswAutorizadores.ColumnHeaders
    .Clear
    .Add , , "Identificación", 2100
    .Add , , "Nombre", 4000
End With

Call sbCargaAutorizadores

End Sub

Private Sub sbCargaAutorizadores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

    'Consulta el Comite del preanalisis
    strSQL = "select isnull(Id_comite,0) as Id_Comite from CRD_PREA_PREANALISIS where COD_PREANALISIS = '" & gPreAnalisis.Expediente & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mComite = rs!ID_COMITE
    Else
        mComite = 0
    End If
    
    rs.Close

    With lswAutorizadores
     .ListItems.Clear
     vPaso = True
     strSQL = "select M.CEDULA,M.NOMBRE,A.CEDULA as 'Asignado'" _
            & " from CRD_COMITES_AUTORIZADORES CA" _
            & " inner join CRD_COMITES_MIEMBROS M on CA.CEDULA = M.CEDULA" _
            & " left join CRD_PREA_AUTORIZADORES A on CA.CEDULA = A.CEDULA" _
            & " and A.COD_PREANALISIS = '" & gPreAnalisis.Expediente _
            & "' where M.ESTADO = 'A'  and CA.ID_COMITE = " & mComite _
            & " order by M.NOMBRE"
            
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , rs!cedula)
          itmX.SubItems(1) = rs!nombre
          If Not IsNull(rs!asignado) Then
             itmX.Checked = vbChecked
             itmX.ForeColor = vbBlue
          End If
      rs.MoveNext
     Loop
     rs.Close
     
     vPaso = False
     
    End With
    
End Sub

Private Sub lswAutorizadores_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

On Error GoTo vError
    
    If Item.Checked Then
      strSQL = "insert CRD_PREA_AUTORIZADORES(COD_PREANALISIS,CEDULA,USUARIO) values('" & gPreAnalisis.Expediente & "','" & Item.Text _
             & "','" & glogon.Usuario & "')"
    Else
      strSQL = "delete CRD_PREA_AUTORIZADORES where CEDULA = '" & Item.Text _
             & "' and COD_PREANALISIS = '" & gPreAnalisis.Expediente & "'"
    End If
    Call ConectionExecute(strSQL)
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

