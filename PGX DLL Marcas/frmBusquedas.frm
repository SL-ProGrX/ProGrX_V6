VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusquedas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busqueda Rápida"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5790
   DrawWidth       =   2
   HelpContextID   =   9004
   Icon            =   "frmBusquedas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfigura 
      Caption         =   "Configuración de las Busquedas"
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox cboBuscar 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cboOrden 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox cbo 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtRegistrosRetorno 
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Text            =   "999"
         ToolTipText     =   "Indicar 999 para que sea Indefinido"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtContadorEventos 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Text            =   "30"
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton cmdDefecto 
         Caption         =   "&Defecto"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar Criterio presionando"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Ordenamiento"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   4080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   120
         X2              =   4080
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Criterio de Busqueda"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Total Registros de Retorno [999 Indefinido]"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Contador de Refrescamiento"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   5715
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2363
            MinWidth        =   2363
            TextSave        =   "03:27 p.m."
            Object.ToolTipText     =   "Contador de Registros"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6773
            MinWidth        =   6773
            Object.ToolTipText     =   "Estado de las Busquedas"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   953
            MinWidth        =   953
            Picture         =   "frmBusquedas.frx":030A
            Object.ToolTipText     =   "Configuración de las Busquedas"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCriterio 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Ingrese aqui el criterio de busqueda"
      Top             =   0
      Width           =   5775
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusquedas.frx":075C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblBusca 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   435
      Width           =   5775
   End
End
Attribute VB_Name = "frmBusquedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub cmdCerrar_Click()
Dim vMensaje As String
vMensaje = ""

If Not IsNumeric(txtContadorEventos) Then vMensaje = vMensaje & " - El Contador de Eventos No es válido..." & vbCrLf
If Not IsNumeric(txtRegistrosRetorno) Then vMensaje = vMensaje & " - El Registro de Retorno No es válido..." & vbCrLf
  
If Len(vMensaje) = 0 Then
  txtCriterio.Enabled = True
  fraConfigura.Visible = False
Else
  MsgBox vMensaje, vbExclamation
End If

End Sub

Private Sub cmdDefecto_Click()

cbo.Text = "01 - Cómo"
cboOrden.Text = "01 - Ascendente"
cboBuscar.Text = "01 - Tecla Retorno"
txtContadorEventos = 30
txtRegistrosRetorno = 999

End Sub

Private Sub cmdGuardar_Click()
Dim fnFile, vMensaje As String

vMensaje = ""

If Not IsNumeric(txtContadorEventos) Then vMensaje = vMensaje & " - El Contador de Eventos No es válido..." & vbCrLf
If Not IsNumeric(txtRegistrosRetorno) Then vMensaje = vMensaje & " - El Registro de Retorno No es válido..." & vbCrLf
  
If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If

fnFile = FreeFile

Open App.Path & "\Busquedas.ini" For Output As #fnFile  ' Create file name.
  Print #fnFile, txtContadorEventos
  Print #fnFile, txtRegistrosRetorno
  Print #fnFile, cboBuscar.Text
  Print #fnFile, cbo.Text
  Print #fnFile, cboOrden.Text
Close #fnFile

MsgBox "Parámetros de Busquedas Guardados Correctamente...", vbInformation

End Sub

Private Sub Form_Load()
Dim fnFile, vArchivo As String
Dim i As Integer

fnFile = FreeFile

lsw.ColumnHeaders.Add , , "", 1200

cbo.Clear
cbo.AddItem "01 - Cómo"
cbo.AddItem "02 - Igual"
cbo.AddItem "03 - Contiene"
cbo.AddItem "04 - Mayor Que"
cbo.AddItem "05 - Menor Que"

cboOrden.Clear
cboOrden.AddItem "01 - Ascendente"
cboOrden.AddItem "02 - Descendente"
 
cboBuscar.Clear
cboBuscar.AddItem "01 - Tecla Retorno"
cboBuscar.AddItem "02 - Tecla F2"
cboBuscar.AddItem "03 - Cualquier Tecla"

'Si no existe el Archivo Crearlo con los valores por defecto
'si existe leer los valores

If Dir(App.Path & "\busquedas.ini", vbArchive) <> "" Then
  i = 1
  Open App.Path & "\Busquedas.ini" For Input As #fnFile  ' Create file name.
  Do While Not EOF(fnFile)
    Input #fnFile, vArchivo
    Select Case i
      Case 1
         txtContadorEventos = vArchivo
      Case 2
         txtRegistrosRetorno = vArchivo
      Case 3
         cboBuscar.Text = Trim(vArchivo)
      Case 4
         cbo.Text = Trim(vArchivo)
      Case 5
         cboOrden.Text = Trim(vArchivo)
      Case Else
    End Select
    i = i + 1
  Loop
  Close #fnFile

Else
    Call cmdDefecto_Click
    Call cmdGuardar_Click
End If
 
 
  
Select Case Mid(cbo.Text, 1, 2)
  Case "01" 'Como
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " cómo [" & Trim(txtCriterio) & "...]"
  Case "02" 'Igual
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " igual a [" & Trim(txtCriterio) & "]"
  Case "03" 'Contiene
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " que contenga [..." & Trim(txtCriterio) & "...]"
  Case "04" 'Mayor Que
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " Mayor que [" & Trim(txtCriterio) & "]"
  Case "05" 'Menor Que
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " Menor que [" & Trim(txtCriterio) & "]"
End Select
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 gBusquedas.Orden = ""
 gBusquedas.Columna = ""
 gBusquedas.Consulta = ""
 gBusquedas.Filtro = ""
 gBusquedas.Convertir = "S"
 gBusquedas.Mascara = ""
End Sub

Private Sub sbBuscar()
Dim strSQL As String

Me.MousePointer = vbHourglass

'No se utiliza TOP en el select, para devolver xx numero de registros
'ya que esta consulta requiere un barrido sobre toda la tabla en ordenamiento
'afectando el rendimiento del foward por el statico, los registros
'se controlaran por medio de un conteo.

If UCase(gBusquedas.Convertir) = "S" Or gBusquedas.Convertir = "" Then
    strSQL = gBusquedas.Consulta & " Where CONVERT(Char(12)," _
           & gBusquedas.Columna & ")"
Else
    strSQL = gBusquedas.Consulta & " Where " & gBusquedas.Columna
End If

Select Case Mid(cbo.Text, 1, 2)
  Case "01" 'Como
     strSQL = strSQL & " like '" & Format(txtCriterio, gBusquedas.Mascara) & "%'"
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " cómo [" & Trim(txtCriterio) & "...]"
  Case "02" 'Igual
     strSQL = strSQL & " = '" & Format(txtCriterio, gBusquedas.Mascara) & "'"
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " igual a [" & Trim(txtCriterio) & "]"
  Case "03" 'Contiene
     strSQL = strSQL & " like '%" & Format(txtCriterio, gBusquedas.Mascara) & "%'"
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " que contenga [..." & Trim(txtCriterio) & "...]"
  Case "04" 'Mayor Que
     strSQL = strSQL & " > '" & Format(txtCriterio, gBusquedas.Mascara) & "'"
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " Mayor que [" & Trim(txtCriterio) & "]"
  Case "05" 'Menor Que
     strSQL = strSQL & " < '" & Format(txtCriterio, gBusquedas.Mascara) & "'"
     lblBusca.Caption = "Busqueda de " & gBusquedas.Columna & " Menor que [" & Trim(txtCriterio) & "]"
End Select
 
lblBusca.Refresh

If Len(Trim(gBusquedas.Filtro)) > 0 Then strSQL = strSQL & " " & gBusquedas.Filtro

If Len(Trim(gBusquedas.Orden)) > 0 Then
  strSQL = strSQL & " Order by " & gBusquedas.Orden
  If Mid(cboOrden.Text, 1, 2) = "02" Then strSQL = strSQL & " desc"
End If

Call sbCargaLsw(strSQL)

Me.MousePointer = vbDefault

End Sub

Private Sub sbCargaLsw(strSQL As String)
Dim i As Integer, itmX As ListItem
Dim x As Integer, y As Integer, IconX As Integer

On Error GoTo vError

lsw.ColumnHeaders.Clear
lsw.ListItems.Clear

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

Status.Panels(1).Style = sbrText
Status.Panels(2) = "Cargando Información..."

If Not rs.BOF And Not rs.EOF Then

    With lsw
     'Carga Titulos
     For i = 0 To (rs.Fields.Count - 1)
        .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
        'Si el Texto de la Columna tiene mas caracteres que los valores
        'Que esta recibe, entonces utilizar el nombre de columna
        If Len(rs.Fields(i).Name) > rs.Fields(i).DefinedSize Then
            .ColumnHeaders.Item(i + 1).Width = Len(rs.Fields(i).Name) * 140
        Else
            .ColumnHeaders.Item(i + 1).Width = rs.Fields(i).DefinedSize * 95
        End If
        If i > 0 Then
            Select Case rs.Fields(i).Type
               Case adCurrency, adDecimal, adDouble, adNumeric
                  .ColumnHeaders.Item(i + 1).Alignment = lvwColumnRight
               Case Else
            End Select
        End If
     Next i
     
     x = 0 'Contador de Refrescamiento
     y = 0 'Contador de Registros
     IconX = 1 'Utiliza el Icon 2 de la lista de imagenes
      Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs.Fields(0).Value, , IconX)
        For i = 1 To (rs.Fields.Count - 1)
            itmX.SubItems(i) = rs.Fields(i).Value & ""
        Next i
        
        If x = Val(txtContadorEventos) Then
          Status.Panels(1) = y
          DoEvents
          x = 0
        Else
          x = x + 1
        End If
        
        y = y + 1
       
        If Val(txtRegistrosRetorno) <> 999 Then
           If y >= Val(txtRegistrosRetorno) Then Exit Do
        End If
      
        rs.MoveNext
      Loop
    End With

End If 'inicio y fin de tabla en true
rs.Close

Status.Panels(2) = "Información Procesada, Total de Registros : " & y
Status.Panels(1).Style = sbrTime

Exit Sub

vError:
 If Err.Number = 3705 Then
    rs.Close
    Call sbCargaLsw(strSQL)
 Else
    MsgBox Err.Description, vbExclamation
 End If

End Sub

Private Sub lsw_Click()
On Error Resume Next

gBusquedas.Resultado = lsw.SelectedItem.Text
gBusquedas.Resultado2 = lsw.SelectedItem.SubItems(1)

Unload Me

End Sub

Private Sub Status_PanelClick(ByVal Panel As MSComctlLib.Panel)
  txtCriterio.Enabled = False
  fraConfigura.Visible = True
End Sub


Private Sub txtCriterio_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case Mid(cboBuscar.Text, 1, 2)
  Case "01" 'Tecla Retorno
    If KeyCode = vbKeyReturn Then Call sbBuscar
  Case "02" 'Tecla Tabulación
    If KeyCode = vbKeyF2 Then Call sbBuscar
  Case "03" 'Cualquier tecla
    Call sbBuscar
End Select
End Sub
