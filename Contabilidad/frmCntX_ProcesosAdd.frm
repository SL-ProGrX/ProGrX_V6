VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCntX_ProcesosAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Procesos Personalizados Adicionales"
   ClientHeight    =   8220
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10692
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5892
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   10572
      _Version        =   1245187
      _ExtentX        =   18648
      _ExtentY        =   10393
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
   Begin XtremeSuiteControls.PushButton cmdProcesar 
      Height          =   612
      Left            =   9000
      TabIndex        =   0
      Top             =   7440
      Width           =   1452
      _Version        =   1245187
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Procesar"
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
      Picture         =   "frmCntX_ProcesosAdd.frx":0000
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   204
      Left            =   120
      TabIndex        =   2
      Top             =   1176
      Width           =   204
      _Version        =   1245187
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      Transparent     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   10572
      _Version        =   1245187
      _ExtentX        =   18648
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Selecciones los Asientos a procesar:"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Procesamiento de Asientos Personalizados!"
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
      Height          =   492
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   8892
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCntX_ProcesosAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdProcesar_Click()
Dim strSQL As String, i As Integer
Dim pProceso As String, pSP_Name As String

Me.MousePointer = vbHourglass

On Error GoTo vError
With lsw.ListItems
    For i = 1 To .Count
      If .Item(i).Checked Then
         pSP_Name = .Item(i).Tag
         strSQL = "exec " & pSP_Name & " '" & .Item(i).Text & "'," & gCntX_Parametros.CodigoConta _
                & "," & gCntX_Parametros.PeriodoAnio & "," & gCntX_Parametros.PeriodoMes _
                & ",'" & glogon.Usuario & "'"
         Call ConectionExecute(strSQL, 0)
         
         Call Bitacora("Aplica", "Proceso Add.:" & .Item(i).Text & " (Conta.:" & gCntX_Parametros.CodigoConta _
                & " Periodo.: " & gCntX_Parametros.PeriodoAnio & "-" & gCntX_Parametros.PeriodoMes)
         
      End If
    Next i
End With

Me.MousePointer = vbDefault

MsgBox "Procesos Adicionales procesados satisfactoriamente!", vbInformation

Unload Me

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Proceso", 1200
    .Add , , "Descripción", 8500
End With

strSQL = "select * from CntX_Procesos_Add" _
       & " where cod_Contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and activo = 1"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Cod_Proceso)
      itmX.SubItems(1) = rs!Descripcion
      itmX.Tag = rs!Sp_Name
  
  rs.MoveNext
Loop
rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

