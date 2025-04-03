VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmAF_RecepcionDevolucionesTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Afiliaciones..: Recepción Devoluciones"
   ClientHeight    =   8148
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11988
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAF_RecepcionDevolucionesTags.frx":0000
   ScaleHeight     =   8148
   ScaleWidth      =   11988
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtCedula 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin MSComctlLib.ListView lswCedula 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   11535
      _ExtentX        =   20341
      _ExtentY        =   10181
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cédula"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Oficina"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Código"
         Object.Width           =   2999
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAplicar 
      Height          =   570
      Left            =   10440
      TabIndex        =   1
      Top             =   7440
      Width           =   1305
      _ExtentX        =   2307
      _ExtentY        =   995
      ButtonWidth     =   2117
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Aplicar Etiqueta"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17166
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   360
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionDevolucionesTags.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionDevolucionesTags.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionDevolucionesTags.frx":13916
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Afiliaciones ..: Recepción de Devoluciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label LblOperacion 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "frmAF_RecepcionDevolucionesTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTagDevolucion As String
Dim mTagAplicado  As String

Private Sub cmdAgregar_Click()
If Trim(txtCedula.Text) <> "" Then Call sbCargaInformacion
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim rs As New ADODB.Recordset

vModulo = 8

  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '11'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagAplicado = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If
    rs.Close



  strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '12'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagDevolucion = Empty
            MsgBox "El código de tag definido el los parámetros para la Recepción/Devolución  no existe"
        End If
        rs.Close
        
    End If

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbAplicarRecepcionDevolucion
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
    End If
End Sub

Private Sub sbCargaInformacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem
On Error GoTo vError

    If Not IsNumeric(txtCedula) Then
        Exit Sub
    End If
    
    If fxValidaNoDuplicados = True Then
        MsgBox "La Cedula se ya fue digitada"
        txtCedula.Text = Empty
        txtCedula.SetFocus
        Exit Sub
    End If
    
    strSQL = "SELECT I.CEDULA,S.nombre,I.CONSEC ,isnull(O.DESCRIPCION,'') as DESCRIPCION " _
            & " FROM AFI_INGRESOS I  inner join SOCIOS S on I.CEDULA = S.CEDULA" _
           & " LEFT JOIN SIF_OFICINAS O ON I.COD_OFICINA = O.COD_OFICINA" _
           & " WHERE I.CEDULA in (select codigo from SIF_CONTROL_TAGS where codigo = '" & Trim(txtCedula.Text) _
           & "' and TAG_CODIGO = 'S04' and cod_modulo = 'AFI') " _
           & " and I.Analista_recepcion = 2 and dbo.fxSIFValidaTagRev(I.cedula,'" & Trim(mTagAplicado) _
           & "','" & Trim(mTagDevolucion) & "','AFI',I.consec,NULL) <> 1"

    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswCedula.ListItems.Add(, , rs!Cedula)
        itmX.SubItems(1) = rs!Nombre
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = rs!Consec
    End If
    rs.Close

    txtCedula.Text = Empty
    txtCedula.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description), vbCritical
        

End Sub

Private Sub sbAplicarRecepcionDevolucion()
Dim i As Integer, strSQL As String

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar esta etiqueta", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If mTagDevolucion = Empty Then
    MsgBox "No se puede realizar el proceso no está definido la etiqueta de devolución"
    Exit Sub
End If


Me.MousePointer = vbHourglass

PrgBar.Max = lswCedula.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswCedula.ListItems

For i = 1 To .Count
        Call sbSIFRegistraTags(.Item(i).Text, mTagDevolucion, "Recepción de Devolución la documentación de la afiliación" _
        , .Item(i).SubItems(3), "AFI", .Item(i).Text, .Item(i).SubItems(3))
        
    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso concluido con éxito...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lswCedula.ListItems.Count

        If Trim(lswCedula.ListItems(i).Text) = Trim(txtCedula.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function




