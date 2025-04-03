VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_SeguimientoPreAnalisisList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de PreAnalisis por Revisar"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Left            =   8400
      Top             =   120
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PreAnalisis"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Operación"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cedula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Seleccione el caso a revisar"
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmCR_SeguimientoPreAnalisisList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Load()
 Call sbInicializa
End Sub


Private Sub lsw_DblClick()
If lsw.ListItems.Count = 0 Then Exit Sub

Operacion.Operacion = lsw.SelectedItem.SubItems(1)
Unload Me

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select P.cod_PreAnalisis,P.id_solicitud,P.cedula,P.nombre,P.fecha_creacion,P.monto" _
       & " from CRD_PREA_PREANALISIS P inner join reg_creditos R on P.id_solicitud = R.id_solicitud and R.estadosol in('R','P')" _
       & " where P.tipo_preAnalisis = 'E' and P.estado = 'A' and P.usuario_gestion = '" & glogon.Usuario _
       & "' and P.id_solicitud is not null"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   Me.MousePointer = vbDefault
   rs.Close
   MsgBox "No Existen Solicitudes por PreAnalisis por Revisar...", vbInformation
   TimerX.Interval = 10
   Exit Sub
   
   
Else
   If rs.RecordCount = 1 Then
        Operacion.Operacion = rs!Id_Solicitud
       
        TimerX.Interval = 10
        Exit Sub
   Else
        Do While Not rs.EOF
          Set itmX = lsw.ListItems.Add(, , rs!cod_PreAnalisis)
              itmX.SubItems(1) = rs!Id_Solicitud
              itmX.SubItems(2) = rs!Cedula
              itmX.SubItems(3) = rs!Nombre
              itmX.SubItems(4) = rs!fecha_creacion
              itmX.SubItems(5) = Format(rs!Monto, "Standard")
          rs.MoveNext
        Loop
   
   End If
End If
rs.Close
 
Me.MousePointer = vbDefault
 
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
Unload Me
End Sub
