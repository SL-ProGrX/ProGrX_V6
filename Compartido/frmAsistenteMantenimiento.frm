VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAsistenteMantenimiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento del Asistente del Sistema"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8535
   Icon            =   "frmAsistenteMantenimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab ssTab 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Asistente"
      TabPicture(0)   =   "frmAsistenteMantenimiento.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdInicializar"
      Tab(0).Control(1)=   "txtForm"
      Tab(0).Control(2)=   "txtTexto"
      Tab(0).Control(3)=   "cmdNuevo"
      Tab(0).Control(4)=   "cbo"
      Tab(0).Control(5)=   "lsw"
      Tab(0).Control(6)=   "cmdGuardar"
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Versiones"
      TabPicture(1)   =   "frmAsistenteMantenimiento.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line2(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label2(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Line2(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboSistema"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "dtpFecha"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtVersion"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtUsuario"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtNotas"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdGuardarVersion"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdGuardarVersion 
         Caption         =   "&Guardar"
         Height          =   435
         Left            =   7200
         TabIndex        =   21
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtNotas 
         Height          =   2355
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   7335
      End
      Begin VB.TextBox txtUsuario 
         Height          =   315
         Left            =   960
         TabIndex        =   18
         Top             =   1200
         Width           =   7335
      End
      Begin VB.TextBox txtVersion 
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   480
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   4440
         TabIndex        =   14
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58458115
         CurrentDate     =   37683
      End
      Begin VB.ComboBox cboSistema 
         Height          =   315
         ItemData        =   "frmAsistenteMantenimiento.frx":0902
         Left            =   960
         List            =   "frmAsistenteMantenimiento.frx":091B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdInicializar 
         Caption         =   "&Inicializar"
         Height          =   315
         Left            =   -69720
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtForm 
         Height          =   315
         Left            =   -73680
         TabIndex        =   5
         Top             =   480
         Width           =   7095
      End
      Begin VB.TextBox txtTexto 
         Height          =   915
         Left            =   -73680
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   7095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   315
         Left            =   -67560
         TabIndex        =   2
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   -73680
         TabIndex        =   1
         Top             =   1800
         Width           =   3615
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   6
         Top             =   2160
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Formulario"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Texto"
            Object.Width           =   11359
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Animación"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   315
         Left            =   -68520
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8280
         X2              =   240
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label2 
         Caption         =   "Notas"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   8280
         X2              =   240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Versión"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Sistema"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Formulario"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Texto"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Animación"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAsistenteMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim vCon As New ADODB.Connection

Private Function fxExiste(vForm As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select formulario from asistente where formulario = '" _
       & vForm & "'"
rs.Open strSQL, vCon, adOpenStatic
If Not rs.BOF And Not rs.EOF Then
   fxExiste = True
Else
   fxExiste = False
End If
rs.Close
End Function


Private Sub cmdGuardar_Click()
Dim strSQL As String, lng As Long, itmX As ListItem

If fxExiste(txtForm) Then
   strSQL = "update asistente set detalle = '" & txtTexto _
          & "',animacion = '" & Trim(cbo.Text) _
          & "' Where formulario = '" & txtForm & "'"
   vCon.Execute strSQL
   For lng = 1 To lsw.ListItems.Count
     If UCase(lsw.ListItems.Item(lng)) = UCase(txtForm) Then
       lsw.ListItems.Item(lng).SubItems(1) = txtTexto
       lsw.ListItems.Item(lng).SubItems(2) = Trim(cbo.Text)
       Exit For
     End If
   Next lng
   
Else

   strSQL = "insert into asistente(formulario,detalle,animacion) values('" & txtForm _
          & "','" & txtTexto & "','" & Trim(cbo.Text) & "')"
   vCon.Execute strSQL
   Set itmX = lsw.ListItems.Add(, , txtForm)
       itmX.SubItems(1) = txtTexto
       itmX.SubItems(2) = Trim(cbo.Text)

End If

txtForm = ""
txtTexto = ""
txtForm.SetFocus

End Sub

Private Sub cmdGuardarVersion_Click()
Dim strSQL As String, lng As Long, itmX As ListItem

On Error GoTo vError

strSQL = "insert into versiones(sistema,version,fecha,idea,descripcion) values('" _
       & cboSistema.Text & "','" & txtVersion & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") _
       & "','" & txtUsuario & "','" & txtNotas & "')"
vCon.Execute strSQL

txtUsuario.SetFocus
txtNotas = ""

MsgBox "Información de la Versión Actualizada...", vbInformation
Exit Sub

vError:
 MsgBox Err.Description, vbCritical


End Sub

Private Sub cmdInicializar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select formulario from opciones group by formulario"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 strSQL = "insert into asistente(formulario,detalle) values('" & Trim(rs!formulario) _
        & "','!')"
 vCon.Execute strSQL
 rs.MoveNext
Loop
rs.Close

Call sbLlenaLsw

End Sub

Private Sub cmdNuevo_Click()
txtForm = ""
txtTexto = ""
txtForm.SetFocus
End Sub


Private Sub Form_Load()
Dim strSQL As String
strSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" _
        & App.Path & "\Asistente.mdb;Mode=ReadWrite;Jet OLEDB:Database Password=PitDead"
vCon.Open strSQL

Call sbLlenaLsw

cbo.Clear
cbo.AddItem ""
For Each AnimationName In Agente.AnimationNames
        cbo.AddItem AnimationName
Next

dtpFecha.Value = fxFechaServidor

End Sub

Private Sub Form_Unload(Cancel As Integer)
vCon.Close
End Sub


Private Sub sbLlenaLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem


Me.MousePointer = vbHourglass

strSQL = "select * from asistente"
rs.Open strSQL, vCon, adOpenForwardOnly

lsw.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!formulario & "")
     itmX.SubItems(1) = rs!detalle & ""
     itmX.SubItems(2) = rs!animacion & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


End Sub

Private Sub lsw_Click()

If lsw.ListItems.Count > 0 Then
   txtForm = lsw.SelectedItem
   txtTexto = lsw.SelectedItem.SubItems(1)
   cbo.Text = lsw.SelectedItem.SubItems(2)
End If

End Sub
