VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmFNDRecepcionDevoluciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos.: Recepción de Devoluciones"
   ClientHeight    =   8016
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11928
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFND_RecepcionDevoluciones.frx":0000
   ScaleHeight     =   8016
   ScaleWidth      =   11928
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtContratoB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
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
      Left            =   7560
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtCedula 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar tlbAplicar 
      Height          =   570
      Left            =   10440
      TabIndex        =   2
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
      Left            =   120
      TabIndex        =   3
      Top             =   7560
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17801
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   240
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
            Picture         =   "frmFND_RecepcionDevoluciones.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionDevoluciones.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionDevoluciones.frx":13916
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lswCedula 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20553
      _ExtentY        =   10393
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
      NumItems        =   6
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
         Text            =   "Operadora"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Plan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contrato"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Plan"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Contrato"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fondos ..: Recepción de Devoluciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label LblOperacion 
      Caption         =   "Cédula"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmFNDRecepcionDevoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTagDevolucion As String
Dim mTagAplicado  As String

Private Sub cmdAgregar_Click()
 Call sbCargaInformacion
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

   strSQL = "SELECT F.CEDULA,S.nombre,F.COD_PLAN,F.COD_OPERADORA,F.COD_CONTRATO" _
          & ",isnull(O.DESCRIPCION,'') as DESCRIPCION" _
          & " FROM FND_CONTRATOS F   inner join SOCIOS S on F.CEDULA = S.CEDULA" _
          & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
          & " WHERE F.cod_Plan = '" & txtCodigo.Text & "'" _
          & "  and F.cod_contrato = '" & txtContratoB.Text & "'" _
          & " and F.Analista_recepcion = 2 " _
          & " "

    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswCedula.ListItems.Add(, , rs!Cedula)
        itmX.SubItems(1) = rs!Nombre
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = rs!Cod_Operadora
        itmX.SubItems(4) = rs!cod_Plan
        itmX.SubItems(5) = rs!cod_contrato
        
    End If
    rs.Close

    txtCedula.Text = Empty
    txtCedula.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

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
        Call sbSIFRegistraTags(.Item(i).SubItems(4), mTagDevolucion, "Recepción de Devolución la documentación del contrato" _
                             , .Item(i).SubItems(5), "FND", .Item(i).SubItems(4), .Item(i).SubItems(5))
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




Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And txtCodigo.Locked = False Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = " And Cod_operadora=1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)

   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
End If
End Sub

Private Sub txtContratoB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbCargaInformacion
End If
End Sub
