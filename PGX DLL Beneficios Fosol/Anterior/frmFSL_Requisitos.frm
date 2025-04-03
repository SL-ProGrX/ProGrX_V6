VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFSL_Requisitos 
   Caption         =   "Requisitos por Causa de Liquidación"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGridRequisitos 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7215
      _Version        =   524288
      _ExtentX        =   12726
      _ExtentY        =   5530
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   485
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_Requisitos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageCombo cboTipo 
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin MSComctlLib.ImageCombo cboCausa 
      Height          =   345
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmFSL_Requisitos.frx":1B90
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Causa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Tipos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos por Causa de Liquidación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmFSL_Requisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vCod_Requisito, vDescripcion, vOpcional, vEstado As String

Private Sub Form_Activate()
 ' vModulo =
End Sub

Private Sub sbCargaTipos()
On Error GoTo error

    strSQL = "select ID_CAUSA,DESCRIPCION  " _
           & "From CAUSAS_RENUNCIAS Where ACTIVO = 1"

    rs.Open strSQL, glogon.Conection, adOpenStatic

    Do While Not rs.EOF
        cboTipo.ComboItems.Add , rs.Fields("ID_CAUSA") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
        rs.MoveNext
    Loop

    rs.Close

Exit Sub

error:
  MsgBox Err.Description
  
End Sub

Private Sub sbCargaCausas()
On Error GoTo error
   
   strSQL = "Select COD_CAUSA, COD_TIPO, DESCRIPCION from FSL_CAUSAS "
          
   rs.Open strSQL, glogon.Conection, adOpenStatic
    
   Do While Not rs.EOF
      cboCausa.ComboItems.Add , rs.Fields("ID_CAUSA") & "(id)", UCase(Trim(rs.Fields("DESCRIPCION")))
      rs.MoveNext
   Loop

   rs.Close
      
Exit Sub

error:
   MsgBox Err.Description


End Sub

Private Sub sbGuardaRequisito()
On Error GoTo error
   
   strSQL = "Insert FSL_REQUISITOS_CAUSAS (COD_REQUISITO, COD_CAUSA, DESCRIPCION, OPCIONAL, ESTADO) " _
          & "Values ('" & vCod_Requisito & "','" & Trim(cboCausa.Text) & "','" & vDescripcion & "', " _
          & "'" & vOpcional & "','" & vEstado & "')"
   glogon.Conection.Execute strSQL
      
Exit Sub

error:
   MsgBox Err.Description
End Sub

Private Sub sbModificaRequisito()
On Error GoTo error
   
  strSQL = "Update FSL_REQUISITOS_CAUSAS set COD_CAUSA='" & Trim(cboCausa.Text) & "', " _
         & "DESCRIPCION= '" & vDescripcion & "',OPCIONAL= '" & vOpcional & "',ESTADO= '" & vEstado & "' " _
         & "where COD_REQUISITO= '" & vCod_Requisito & "'"
  
  glogon.Conection.Execute strSQL
          
Exit Sub

error:
   MsgBox Err.Description

End Sub

Private Sub sbBorraRequisito()
On Error GoTo error
   
   strSQL = "Delete FSL_REQUISITOS_CAUSAS where COD_REQUISITO= '" & vCod_Requisito & "'"
   
   glogon.Conection.Execute strSQL
       
Exit Sub

error:
   MsgBox Err.Description
   
End Sub

Private Function fxValidaRequisito() As Boolean
Dim I As Integer

On Error GoTo error
   
   With vGridRequisitos
   
      .Row = .ActiveRow
      
      .Col = 1
      vCod_Requisito = .Text
      
      strSQL = "Select COD_REQUISITO from dbo.FSL_REQUISITOS_CAUSAS"
      rs.Open strSQL, glogon.Conection, adOpenStatic

      If vCod_Requisito = rs.Fields("COD_REQUISITO") Then
         fxValidaRequisito = True
      Else
         fxValidaRequisito = False
      End If
        
      rs.Close
        
      .Col = 2
      vDescripcion = .Text
        
      .Col = 3
      vOpcional = .Text
      
      .Col = 4
      vEstado = .Text
      
   End With
   
Exit Function

error:
  MsgBox Err.Description
  
End Function

