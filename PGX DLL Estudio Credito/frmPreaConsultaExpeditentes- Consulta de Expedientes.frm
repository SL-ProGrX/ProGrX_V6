VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPreaConsultaExpeditentes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Expedientes y Subexpedientes"
   ClientHeight    =   5520
   ClientLeft      =   -1845
   ClientTop       =   4425
   ClientWidth     =   9030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView ltvExpedientes 
      Height          =   4185
      Left            =   0
      TabIndex        =   7
      Top             =   1050
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7382
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Expediente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sub Expediente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   5380
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "C�dula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha Creaci�n"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   8955
      Begin VB.TextBox TxtExpediente 
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "N�mero de expediente"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtCedula 
         Height          =   315
         Left            =   6900
         TabIndex        =   3
         ToolTipText     =   "C�dula de identidad"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1950
         TabIndex        =   2
         ToolTipText     =   "Nombre"
         Top             =   480
         Width           =   4965
      End
      Begin VB.Label Label2 
         Caption         =   "N� Expediente"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1950
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "C�dula"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6900
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   450
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPreaConsultaExpeditentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private clsEntidad As New ASE_PreAnalisis.clsEntidad
Private clsNull As New ASE_PreAnalisis.clsNull
Public m_Expediente As String
Private ItemSeleccionado As MSComctlLib.ListItem
Private vLItem As MSComctlLib.ListItem

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn) Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Call gsbCentrarModales(Me)
End Sub
Public Function fxAgregaColleccion(ByVal pExpediente As String, ByVal pNombre As String, ByVal pCedula As String) As String
On Error GoTo Error
Dim Vcoleccion As New Collection
With Vcoleccion
    .Add fxFormatearValor(pExpediente, Caracter)
    .Add fxFormatearValor(pNombre, Caracter)
    .Add fxFormatearValor(pCedula, Caracter)
End With
fxAgregaColleccion = fxFormatearValuesCollection(Vcoleccion)

Exit Function
Error:
    MsgBox Err.Description
End Function
Private Sub BuscarExpediente()

On Error GoTo Error
m_Expediente = ""
glogon.strSQL = ""
If (Len(Trim(txtExpediente.Text)) = 0) And (Len(Trim(txtNombre.Text)) = 0) And (Len(Trim(txtCedula.Text)) = 0) Then
   ' MsgBox "La consulta que intenta realizar es una consulta muy general. Es necesario que introduzca por lo menos un criterio para la b�squeda.", vbInformation, gTituloMsg
    txtExpediente.SetFocus
    ltvExpedientes.ListItems.Clear
    Exit Sub
End If
    
    Screen.MousePointer = vbHourglass
    ltvExpedientes.ListItems.Clear
    clsEntidad.tablaName = "spCRDPreaPREANALISIS"
    If clsEntidad.fxTraerFiltrado("Criterios", fxAgregaColleccion(txtExpediente.Text & "%", "%" & txtNombre.Text & "%", txtCedula.Text & "%")) Then
        Call sbCargaLista
    Else
        MsgBox "No se encontr� ning�n asegurado con los criterios especificados.", vbInformation, gTituloMsg
    End If
    Screen.MousePointer = vbDefault
    If (ltvExpedientes.ListItems.Count > 0) Then
'        ltvExpedientes.SetFocus
        ltvExpedientes.ListItems(1).Selected = True
    End If
    
salir:
    Exit Sub
Error:
    Screen.MousePointer = vbDefault
    cMensaje.deError ("Ocurri� un error consutaldo expedientes.")
End Sub

Private Sub sbCargaLista()
On Error GoTo VError
    
    Dim icono As Integer
    
    Screen.MousePointer = vbHourglass
    
    ProgressBar.Visible = True
    ProgressBar.Left = ltvExpedientes.Left
    ProgressBar.Width = ltvExpedientes.Width
    ltvExpedientes.Height = ltvExpedientes.Height - 315
    ProgressBar.Top = ltvExpedientes.Top + ltvExpedientes.Height
    
    ltvExpedientes.ListItems.Clear
    DoEvents
    ProgressBar.Value = 0
    ProgressBar.Max = glogon.Recordset.RecordCount
    
With glogon.Recordset
     While Not glogon.Recordset.EOF

        ProgressBar.Value = ProgressBar.Value + 1
        
        Set vLItem = ltvExpedientes.ListItems.Add(, .Fields("COD_PREANALISIS") & "id", .Fields("COD_PREANALISIS"))
        vLItem.SubItems(1) = .Fields("COD_PREANALISIS_REF")
        vLItem.SubItems(2) = .Fields("NOMBRE")
        vLItem.SubItems(3) = .Fields("CEDULA")
        vLItem.SubItems(4) = Format(.Fields("FECHA_CREACION"), "dd/mm/yyyy")
        vLItem.SubItems(5) = .Fields("USUARIO")
       .MoveNext
    Wend
End With

Screen.MousePointer = vbDefault
    
salir:
    Screen.MousePointer = vbDefault
    ProgressBar.Visible = False
    ltvExpedientes.Height = ltvExpedientes.Height + 315
    Exit Sub
VError:
    cMensaje.deError ("ocurrio un error al cargar las lista")
    
    Resume salir
    
End Sub

Private Sub ltvExpedientes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 ltvExpedientes.SortKey = ColumnHeader.Index - 1
    
    If (ltvExpedientes.SortOrder = lvwAscending) Then
        ltvExpedientes.SortOrder = lvwDescending
    Else
        ltvExpedientes.SortOrder = lvwAscending
    End If
    
    ltvExpedientes.Sorted = True
End Sub

Private Sub ltvExpedientes_DblClick()
If Not ItemSeleccionado Is Nothing Then
    m_Expediente = ltvExpedientes.SelectedItem.Text
    Unload Me
End If
 
End Sub

Private Sub ltvExpedientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Error
    Set ItemSeleccionado = Item
    txtExpediente.Text = Item.Text
    txtNombre.Text = Item.SubItems(2)
    txtCedula.Text = Item.SubItems(3)
Exit Sub
Error:
  MsgBox Err.Description, vbExclamation
End Sub


Private Sub txtCedula_Validate(Cancel As Boolean)
    Call BuscarExpediente
End Sub

Private Sub TxtExpediente_Validate(Cancel As Boolean)
    Call BuscarExpediente
End Sub

Private Sub txtNombre_Validate(Cancel As Boolean)
    Call BuscarExpediente
End Sub
