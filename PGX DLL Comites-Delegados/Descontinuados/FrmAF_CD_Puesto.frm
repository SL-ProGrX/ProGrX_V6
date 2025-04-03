VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmAF_CD_Puesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimineto de Puestos"
   ClientHeight    =   4515
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7665
   Icon            =   "FrmAF_CD_Puesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAplicar 
      Height          =   855
      Left            =   6645
      Picture         =   "FrmAF_CD_Puesto.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Aplicar puestos"
      Top             =   210
      Width           =   825
   End
   Begin FPSpreadADO.fpSpread vGridPuestos 
      Height          =   3045
      Left            =   255
      TabIndex        =   2
      Top             =   1200
      Width           =   7230
      _Version        =   458752
      _ExtentX        =   12753
      _ExtentY        =   5371
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "FrmAF_CD_Puesto.frx":398D
      VScrollSpecialType=   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Puestos para los Comites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   3660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Puestos para los Comites"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   255
      TabIndex        =   1
      Top             =   375
      Width           =   3660
   End
End
Attribute VB_Name = "FrmAF_CD_puesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub sbGuarda()

Dim strSql As String, rs As New ADODB.Recordset
Dim A As Integer, B As Integer
Dim InsInfo() As String
ReDim InsInfo(vGridPuestos.MaxCols)


For A = 1 To vGridPuestos.MaxRows
    vGridPuestos.Row = A
        
        For B = 1 To vGridPuestos.MaxCols
          vGridPuestos.Col = B
          InsInfo(B) = vGridPuestos.Text
        Next B
    
    If InsInfo(1) = "" Then
     MsgBox "No hay datos para el código de puesto", vbInformation, "Información"
     Exit Sub
    End If
     
   strSql = "select cod_puesto from afi_cd_puestos " _
            & " where cod_puesto ='" & InsInfo(1) & "'"
            rs.Open strSql, glogon.Conection, adOpenStatic
   
       If rs.EOF Then
             strSql = "insert into afi_cd_puestos (cod_puesto,descripcion,activo)" _
                    & "values('" & UCase(InsInfo(1)) & "','" & UCase(InsInfo(2)) & "'," & InsInfo(3) & ")"
                
         Else
           strSql = "update afi_cd_puestos set " _
                    & "descripcion ='" & UCase(InsInfo(2)) & "'," _
                    & "activo='" & InsInfo(3) & "'" _
                    & "where cod_puesto ='" & InsInfo(1) & "' "
       End If
glogon.Conection.Execute strSql
rs.Close
Next A
MsgBox "Se ingresaron correctamente los datos", vbInformation + vbOKOnly, "Información"
End Sub

Private Sub CmdAplicar_Click()
 Call sbGuarda
End Sub
Sub SbPuestos()

Dim rs As New ADODB.Recordset
Dim strSql, Tipo As String
Dim A, B  As Integer

 strSql = "select cod_puesto,descripcion,activo from afi_cd_puestos "
           rs.Open strSql, glogon.Conection, adOpenStatic
           
vGridPuestos.MaxRows = rs.RecordCount

While Not rs.EOF
  For A = 1 To vGridPuestos.MaxRows
    vGridPuestos.Row = A
       For B = 1 To 3
         vGridPuestos.Col = B
            Select Case True
              Case vGridPuestos.Col = 1
               vGridPuestos.Text = rs!cod_puesto
              Case vGridPuestos.Col = 2
               vGridPuestos.Text = rs!Descripcion
              Case vGridPuestos.Col = 3
               vGridPuestos.Text = rs!activo
              End Select
       Next B
  rs.MoveNext
  Next A
Wend
rs.Close
End Sub
Private Sub Form_Load()
  Call SbPuestos
End Sub


Private Sub vGridpuestos_KeyDown(KeyCode As Integer, Shift As Integer)
 'Inserta Linea
 If KeyCode = vbKeyInsert Then
    vGridPuestos.MaxRows = vGridPuestos.MaxRows + 1
    vGridPuestos.InsertRows vGridPuestos.ActiveRow, 1
    vGridPuestos.Row = vGridPuestos.ActiveRow
 End If
End Sub
