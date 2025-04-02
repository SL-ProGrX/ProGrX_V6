VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCntX_UtilEliminaConta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Eliminación de Contabilidades"
   ClientHeight    =   6012
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   9180
   HelpContextID   =   1
   Icon            =   "frmCntX_UtilEliminaConta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6012
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3132
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   8652
      _Version        =   1245187
      _ExtentX        =   15261
      _ExtentY        =   5524
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.TextBox txtConfirmacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin XtremeSuiteControls.PushButton cmdEliminar 
      Height          =   696
      Left            =   6480
      TabIndex        =   5
      Top             =   4920
      Width           =   2412
      _Version        =   1245187
      _ExtentX        =   4254
      _ExtentY        =   1235
      _StockProps     =   79
      Caption         =   "Eliminar Contabilidad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCntX_UtilEliminaConta.frx":030A
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Eliminación de Contabilidades"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   7692
   End
   Begin VB.Image imgCodigo 
      Appearance      =   0  'Flat
      Height          =   384
      Left            =   5040
      Picture         =   "frmCntX_UtilEliminaConta.frx":0AD7
      Top             =   4920
      Width           =   384
   End
   Begin VB.Label lblConfirmacion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   312
      Left            =   3600
      TabIndex        =   3
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Digite el código de confirmación >"
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
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   5280
      Width           =   3492
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Confirmación"
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
      Left            =   360
      TabIndex        =   1
      Top             =   4920
      Width           =   3492
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleccione las Contabilidades No Deseadas Para Eliminarlas de la Base de Datos"
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
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8652
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1104
      Left            =   0
      Top             =   0
      Width           =   10344
   End
End
Attribute VB_Name = "frmCntX_UtilEliminaConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEliminar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, i As Integer, lng As Long

On Error GoTo vError

If lblConfirmacion.Caption <> txtConfirmacion.Text Then
  MsgBox "Código de Confirmación no es válido revise...!", vbExclamation
  Exit Sub
End If

i = MsgBox("Esta Seguro de Eliminar las Contabilidades Marcadas, este afectará sus consolidaciones " _
          & "en donde esta contabilidad es miembro..", vbYesNo)

If i = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

frmCntX_Procesos.Show
frmCntX_Procesos.Caption = "Eliminando Contabilidades"
frmCntX_Procesos.Refresh
frmCntX_Procesos.TimerX.Interval = 800

With lsw.ListItems

For lng = 1 To .Count
 If .Item(lng).Checked Then
    
  frmCntX_Procesos.lbl.Caption = "Eliminando Contabilidad: " & .Item(lng).SubItems(1)
  frmCntX_Procesos.lbl.Refresh
    
  strSQL = "exec spCntX_Util_Contabilidad_Elimina " & .Item(lng).Text & ",'" & glogon.Usuario & "','*xHM1tOk3n$'"
    
  Call ConectionExecute(strSQL)
    
  Call Bitacora("Elimina", "Contabilidad: [" & .Item(lng).Text & "] " & .Item(lng).SubItems(1))
  
 End If
Next lng

End With

UnLoad frmCntX_Procesos

Me.MousePointer = vbDefault

MsgBox "Proceso Concluido satisfactoriamente, se recomienda utilizar " _
     & "el Programa de Mantenimiento", vbInformation

Call Form_Load

Exit Sub

vError:
 Me.MousePointer = vbDefault
 frmCntX_Procesos.Hide
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 20

On Error GoTo vError


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1100
    .Add , , "Descripción", 7450
End With
lsw.Checkboxes = True
lsw.ListItems.Clear


strSQL = "Select COD_CONTABILIDAD,NOMBRE from CntX_Contabilidades"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_CONTABILIDAD)
      itmX.SubItems(1) = rs!Nombre
  rs.MoveNext
Loop
rs.Close


Call Formularios(Me)
Call RefrescaTags(Me)

Call imgCodigo_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub imgCodigo_Click()
Dim x As Integer, y As Integer

lblConfirmacion.Caption = ""
txtConfirmacion.Text = ""

For x = 1 To 7
    Randomize
    y = 0
    Do While y < 48
        y = CInt(126 * Rnd())
    Loop
    lblConfirmacion.Caption = lblConfirmacion.Caption & Chr(y)
Next x

End Sub
