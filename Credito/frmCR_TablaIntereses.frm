VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCR_TablaIntereses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Intereses"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   HelpContextID   =   3028
   Icon            =   "frmCR_TablaIntereses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lswDetalle 
      Height          =   2655
      Left            =   0
      TabIndex        =   27
      Top             =   2880
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4683
      View            =   3
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Monto Inicial"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Monto Final"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Int.Cor.Soc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Int.Mor.Soc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Int.Cor.NSoc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Int.Mor.NSoc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Plazo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraActualizaIntereses 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdActualizarIntereses 
         Caption         =   "Actualiza préstamos"
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "Aplicar todos los Casos"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   21
         ToolTipText     =   "Casos Existentes y Nuevos"
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton optActualiza 
         Caption         =   "Aplicar a los Nuevos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   1005
      ButtonWidth     =   1561
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Insertar"
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Borrar"
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Deshacer"
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Ayuda"
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cerrar"
            Key             =   "cerrar"
            Object.ToolTipText     =   "cierra esta ventana"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraIntereses 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   6775
      Begin VB.TextBox txtPlazo 
         Height          =   315
         Left            =   840
         MaxLength       =   3
         TabIndex        =   24
         ToolTipText     =   "Plazo recomendado en este rango"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtIntMorSocios 
         Height          =   315
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   4
         ToolTipText     =   "Intereses corrientes para socios"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtIntCorSocios 
         Height          =   315
         Left            =   4320
         MaxLength       =   3
         TabIndex        =   3
         ToolTipText     =   "Intereses corrientes para socios"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtIntCorNoSocios 
         Height          =   315
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   5
         ToolTipText     =   "Intereses corrientes para socios"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtIntMorNoSocios 
         Height          =   315
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   6
         ToolTipText     =   "Intereses corrientes para socios"
         Top             =   840
         Width           =   615
      End
      Begin MSMask.MaskEdBox medDe 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medHasta 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###,###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERESES  -  SOCIOS      NO SOCIOS"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meses"
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   25
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Plazo"
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONTOS"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "De"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Corrientes"
         Height          =   165
         Index           =   0
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Moratorios"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6000
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   8
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   7
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rangos de Plazos e Intereses Establecidos"
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Top             =   2640
      Width           =   6855
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   6855
   End
End
Attribute VB_Name = "frmCR_TablaIntereses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intEdita As Integer, rsRangos As New ADODB.Recordset
Dim strCodigo As String, Columna As Long 'Almacena la columna llave de rangos
Dim strActualizaSocios As String 'Se utiliza para indicar si se aplica el cambio de intereses
                                 'A todos los socios o exsocios, etc.
Dim intCambiaIntereses As Integer 'se utiliza como control para que el usuario no pueda salir
                                 'con X del form cuando este cambiando intereses
Private Type Inter
  InteresCorrienteSocios As Integer
  InteresMoratorioSocios As Integer
  InteresCorrienteNoSocios As Integer
  InteresMoratorioNoSocios As Integer
  MontoDe As Currency
  montohasta As Currency
End Type
Dim Interes As Inter
Private miTran As Integer
Sub LimpiaPantalla()
    
medDe = ""
medHasta = ""
txtIntCorSocios = ""
txtIntMorSocios = ""
txtIntCorNoSocios = ""
txtIntMorNoSocios = ""
txtPlazo = ""

End Sub
Sub habilita(str As String)
If str = "S" Then
    medDe.Enabled = True
    medHasta.Enabled = True
    txtIntCorSocios.Enabled = True
    txtIntMorSocios.Enabled = True
    txtIntCorNoSocios.Enabled = True
    txtIntMorNoSocios.Enabled = True
    txtPlazo.Enabled = True
Else
    medDe.Enabled = False
    medHasta.Enabled = False
    txtIntCorSocios.Enabled = False
    txtIntMorSocios.Enabled = False
    txtIntCorNoSocios.Enabled = False
    txtIntMorNoSocios.Enabled = False
    txtPlazo.Enabled = False
End If
End Sub
Private Sub cmdActualizarIntereses_Click()
Dim strSQL As String

'NO HACE FALTA CONTROL DE ERRORES PUES YA LO TIENE EN EN UNA
'FUNCION QUE LLAMA A ESTA.

Me.MousePointer = vbHourglass

Select Case True
 Case optActualiza(0).Value 'Solo aplica a los prestamos nuevos
   'No debe de hacer nada
 Case optActualiza(1).Value 'Aplica a los prestamos anteriores y nuevos
    'Tener cuidado con los saldos en cero, cancelados y cobro judicial
    'Los intereses a morosos no se actualiza pues se calcula mensualmente
    'Que pasa con los no socios? son opex o que
    'PARA CONTINUAR EL PROCEDIMIENTO LOS CASOS DE NO SOCIOS SE TRABAJAN COMO OPEX
    'Tiene que verificar que el monto aprobado este dentro del rango
    
    strSQL = "update reg_creditos set interesv = " & txtIntCorSocios.Text
    Select Case strActualizaSocios
      Case "S"
         strSQL = strSQL & " where saldo > 0 and estado = 'A' and proceso = 'N' and opex = 0"
      Case "E"
         strSQL = strSQL & " where saldo > 0 and estado = 'A' and proceso = 'N' and opex = 1"
      Case "T"
         strSQL = strSQL & " where saldo > 0 and estado = 'A' and proceso = 'N'"
    End Select
    strSQL = strSQL + " and montoapr > " & medDe.Text & " and montoapr < " & medHasta.Text
    glogon.Conection.Execute strSQL
    Call Bitacora("Modifica", "Prestamos Intereses Codigo " & frmCR_CatalogoCreditos.txtCodigoCorriente)
    
End Select

If miTran = 1 Then
glogon.Conection.CommitTrans
miTran = 0
End If

intCambiaIntereses = 0

tlbPrincipal.Enabled = True
fraActualizaIntereses.Visible = False
Me.Height = Me.Height - fraActualizaIntereses.Height

Me.MousePointer = vbDefault

End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()

vModulo = 3
Call Formularios(Me)

Call sbToolBar_Iconos(tlbPrincipal)

tlbPrincipal.Buttons.Item(1).Enabled = True 'insertar
tlbPrincipal.Buttons.Item(2).Enabled = False 'modificar
tlbPrincipal.Buttons.Item(3).Enabled = False 'borrar
tlbPrincipal.Buttons.Item(4).Enabled = False 'guardar
tlbPrincipal.Buttons.Item(5).Enabled = False  'deshacer
tlbPrincipal.Buttons.Item(6).Enabled = True  'Ayuda
tlbPrincipal.Buttons.Item(7).Enabled = True  'cerrar

Columna = 0
intEdita = 2
strCodigo = frmCR_CatalogoCreditos.txtCodigoCorriente.Text

lblTitulo.Caption = "Código [" + strCodigo + "]   Descripción [" + frmCR_CatalogoCreditos.txtDescripcion.Text & "]"
lblTitulo.Refresh

Call CargaGrid
Call habilita("N")
Call RefrescaTags(Me)

End Sub
Private Sub Form_Unload(Cancel As Integer)
 If intCambiaIntereses = 1 Then Cancel = vbCancel
End Sub


Private Sub lswDetalle_Click()
On Error Resume Next
With lswDetalle.SelectedItem
 Columna = .Text
 medDe.Text = .SubItems(1)
 medHasta.Text = .SubItems(2)
 txtIntCorSocios.Text = .SubItems(3)
 txtIntMorSocios.Text = .SubItems(4)
 txtIntCorNoSocios.Text = .SubItems(5)
 txtIntMorNoSocios.Text = .SubItems(6)
 txtPlazo.Text = .SubItems(7)
'Mete los valores en la variable de intereses para posible actualización de intereses
'en REG_CREDITOS

 Interes.InteresCorrienteSocios = .SubItems(3)
 Interes.InteresCorrienteNoSocios = .SubItems(5)
 Interes.InteresMoratorioSocios = .SubItems(4)
 Interes.InteresMoratorioNoSocios = .SubItems(6)
 Interes.MontoDe = .SubItems(1)
 Interes.montohasta = .SubItems(2)
 
End With
 
 tlbPrincipal.Buttons.Item(1).Enabled = True 'insertar
 tlbPrincipal.Buttons.Item(2).Enabled = True 'modificar
 tlbPrincipal.Buttons.Item(3).Enabled = True 'borrar
 tlbPrincipal.Buttons.Item(4).Enabled = False 'guardar
 tlbPrincipal.Buttons.Item(5).Enabled = False  'deshacer
 tlbPrincipal.Buttons.Item(6).Enabled = True  'Ayuda
 tlbPrincipal.Buttons.Item(7).Enabled = True  'cerrar

End Sub


Private Sub medDe_Change()
'Set GLOBALES.gCajaTxt = medDe.Text
'ValidaMonto

End Sub

Private Sub medDe_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then medHasta.SetFocus
End Sub
Private Sub medHasta_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtIntCorSocios.SetFocus
End Sub
Private Sub txtIntCorSocios_KeyPress(KeyAscii As Integer)
Call Valida(KeyAscii)
 If KeyAscii = vbKeyReturn Then
   If Val(txtIntCorSocios.Text) > 0 And Val(txtIntCorSocios.Text) < 101 Then txtIntMorSocios.SetFocus
 End If
End Sub

Private Sub txtIntMorNoSocios_KeyPress(KeyAscii As Integer)
Call Valida(KeyAscii)
End Sub

Private Sub txtIntMorSocios_KeyPress(KeyAscii As Integer)
Call Valida(KeyAscii)
 If KeyAscii = vbKeyReturn Then
   If Val(txtIntMorSocios.Text) > 0 And Val(txtIntMorSocios.Text) < 101 Then txtIntCorNoSocios.SetFocus
 End If
End Sub
Private Sub txtIntCorNoSocios_KeyPress(KeyAscii As Integer)
Call Valida(KeyAscii)
 If KeyAscii = vbKeyReturn Then
   If Val(txtIntCorNoSocios.Text) > 0 And Val(txtIntCorNoSocios.Text) < 101 Then txtIntMorNoSocios.SetFocus
 End If
End Sub
Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

If Button.Key <> "cerrar" Then
Me.MousePointer = vbHourglass
End If


Select Case Button.Key
  Case "insertar"
         intEdita = 0
         tlbPrincipal.Buttons.Item(1).Enabled = False 'insertar
         tlbPrincipal.Buttons.Item(2).Enabled = False 'modificar
         tlbPrincipal.Buttons.Item(3).Enabled = False 'borrar
         tlbPrincipal.Buttons.Item(4).Enabled = True 'salvar
         tlbPrincipal.Buttons.Item(5).Enabled = True  'deshacer
         tlbPrincipal.Buttons.Item(6).Enabled = False  'Ayuda
         tlbPrincipal.Buttons.Item(7).Enabled = False  'cerrar
         Call habilita("S")
         LimpiaPantalla
         
  Case "modificar"
         If Columna > 0 Then
          intEdita = 1
          tlbPrincipal.Buttons.Item(1).Enabled = False 'insertar
          tlbPrincipal.Buttons.Item(2).Enabled = False 'modificar
          tlbPrincipal.Buttons.Item(3).Enabled = False 'borrar
          tlbPrincipal.Buttons.Item(4).Enabled = True  'guardar
          tlbPrincipal.Buttons.Item(5).Enabled = True  'deshacer
          tlbPrincipal.Buttons.Item(6).Enabled = False 'Ayuda
          tlbPrincipal.Buttons.Item(7).Enabled = False 'salir
          Call habilita("S")
         Else
          MsgBox "Seleccione un rango", vbOKOnly
          Exit Sub
         End If
         
  Case "borrar"
       'aqui borra y luego carga
      If Columna > 0 Then
        If MsgBox("Está seguro que desea borrar este rango", vbYesNo) = vbYes Then
          glogon.Conection.Execute "Delete rangos where consec = " & Columna
          Call Bitacora("Borra", "Rangos consec=" & Trim(Columna))
          Call LimpiaPantalla
          Call CargaGrid
        End If
      End If
      Call RefrescaTags(Me)
  Case "guardar"
         If ValidaExistencia And intEdita <> 2 Then  'existen todos los datos de la pantalla
            
            tlbPrincipal.Buttons.Item(1).Enabled = True   'insertar
            tlbPrincipal.Buttons.Item(2).Enabled = True   'modificar
            tlbPrincipal.Buttons.Item(3).Enabled = True   'borrar
            tlbPrincipal.Buttons.Item(4).Enabled = False  'guardar
            tlbPrincipal.Buttons.Item(5).Enabled = False  'deshacer
            tlbPrincipal.Buttons.Item(6).Enabled = True   'ayuda
            tlbPrincipal.Buttons.Item(7).Enabled = True   'cerrar
            
            Call Guardar
            
            Call CargaGrid
            intEdita = 0
            Call habilita("N")
         Else
            MsgBox "No se puede guardar el rango especificado, verifique la información", vbCritical
         End If
         Call RefrescaTags(Me)
         
  Case "deshacer"
         Call LimpiaPantalla
         tlbPrincipal.Buttons.Item(1).Enabled = True  'insertar
         tlbPrincipal.Buttons.Item(2).Enabled = True 'modificar
         tlbPrincipal.Buttons.Item(3).Enabled = True  'borrar
         tlbPrincipal.Buttons.Item(4).Enabled = False 'salvar
         tlbPrincipal.Buttons.Item(5).Enabled = False 'deshacer
         tlbPrincipal.Buttons.Item(6).Enabled = True  'ayuda
         tlbPrincipal.Buttons.Item(7).Enabled = True  'salir
         Call RefrescaTags(Me)
  Case "ayuda"
        MDIPrincipal.dlg.HelpContext = Me.HelpContextID
        MDIPrincipal.dlg.ShowHelp
  Case "cerrar"
   
    Unload Me

End Select
If Button.Key <> "cerrar" Then
Me.MousePointer = vbDefault
End If

End Sub

Sub CargaGrid()
Dim itmX As ListItem, rs As New ADODB.Recordset

On Error GoTo CapturaError
rs.Source = "Select consec,de,hasta,intc_soc,intm_soc,intc_nsoc," _
         & "intm_nsoc,plazo from rangos where codigo = '" & strCodigo & "' Order by de"
rs.Open , glogon.Conection, adOpenStatic

With lswDetalle
  .ListItems.Clear
  Do While rs.EOF = False
    Set itmX = .ListItems.Add(.ListItems.Count + 1, , rs!consec)
        itmX.Tag = itmX.Index
        itmX.SubItems(1) = Format(rs!de, "###,###,###,##0.00")
        itmX.SubItems(2) = Format(rs!hasta, "###,###,###,##0.00")
        itmX.SubItems(3) = rs!intc_soc
        itmX.SubItems(4) = rs!intm_soc
        itmX.SubItems(5) = rs!intc_nsoc
        itmX.SubItems(6) = rs!intm_nsoc
        itmX.SubItems(7) = rs!Plazo
    rs.MoveNext
  Loop
End With

rs.Close

Exit Sub
CapturaError:
Call ProcedimientoErrores(Me.Name, Err)

End Sub

Function ValidaExistencia() As Boolean
Dim rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo CapturaError

ValidaExistencia = True
i = 1


'verifica que las casillas sean numericas
If Not IsNumeric(medDe.Text) Then
  ValidaExistencia = False
  i = 0
End If
If Not IsNumeric(medHasta.Text) Then
  ValidaExistencia = False
  i = 0
End If
If Not IsNumeric(txtIntCorSocios.Text) Then
  ValidaExistencia = False
  i = 0
End If
If Not IsNumeric(txtIntMorSocios.Text) Then
  ValidaExistencia = False
  i = 0
End If
If Not IsNumeric(txtIntCorNoSocios.Text) Then
  ValidaExistencia = False
  i = 0
End If
If Not IsNumeric(txtIntMorNoSocios.Text) Then
  ValidaExistencia = False
  i = 0
End If

If Not IsNumeric(txtPlazo.Text) Then
  ValidaExistencia = False
  i = 0
End If


'verifica los montos en pantalla

Select Case ""
 Case Trim(medDe.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(medHasta.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(txtIntCorSocios.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(txtIntMorSocios.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(txtIntCorNoSocios.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(txtIntMorNoSocios.Text)
  ValidaExistencia = False
  i = 0
 Case Trim(txtPlazo.Text)
  ValidaExistencia = False
  i = 0
End Select

If Val(medDe.Text) > Val(medHasta.Text) Then
  ValidaExistencia = False
  i = 0
End If
'Valida nuevamente si los porcentajes esta de 1 a 100 en los intereses

If Val(txtIntCorSocios.Text) < 0 And Val(txtIntCorSocios.Text) > 101 Then
  ValidaExistencia = False
  i = 0
End If
If Val(txtIntMorSocios.Text) < 0 And Val(txtIntMorSocios.Text) > 101 Then
  ValidaExistencia = False
  i = 0
End If
If Val(txtIntCorNoSocios.Text) < 0 And Val(txtIntCorNoSocios.Text) > 101 Then
  ValidaExistencia = False
  i = 0
End If
If Val(txtIntMorNoSocios.Text) < 0 And Val(txtIntMorNoSocios.Text) > 101 Then
  ValidaExistencia = False
  i = 0
End If

If Val(txtPlazo.Text) <= 0 Then
  ValidaExistencia = False
  i = 0
End If


'Verifica si el monto en rangos sea valido o no exista en la BD
If i = 1 Then
With rs
 .Source = "select * from rangos where codigo = '" & strCodigo & "'"
 .ActiveConnection = glogon.Conection
 .CursorType = adOpenStatic
 .Open
 If .EOF = True And .BOF = True Then
  'nada
 Else
 
  Do While .EOF = False
   Select Case Val(medDe.Text)
     Case Is = !de
        If !consec <> Columna Then ValidaExistencia = False
     Case Is < !de
        If !consec <> Columna And Val(medHasta.Text) > !hasta Then ValidaExistencia = False
        If !consec <> Columna And Val(medHasta.Text) >= !de Then ValidaExistencia = False
     Case Is > !de
        If !consec <> Columna And Val(medHasta.Text) < !hasta Then ValidaExistencia = False
        If !consec <> Columna And Val(medDe.Text) <= !hasta Then ValidaExistencia = False
     Case Is = !hasta
        If !consec <> Columna Then ValidaExistencia = False
   End Select
   .MoveNext
  Loop
 End If
 
 .Close
 
End With
End If 'del integer

Exit Function
CapturaError:
Call ProcedimientoErrores(Me.Name, Err)

End Function


Sub Guardar()
'Inserta o actualiza la información de rangos
Dim strSQL, rs As New ADODB.Recordset

If miTran = 0 Then
glogon.Conection.BeginTrans
miTran = 1
End If

If ValidaExistencia Then

Select Case intEdita
 Case 0 'inserta
   strSQL = "insert into rangos(codigo,de,hasta,intm_soc,intc_soc,intm_nsoc,intc_nsoc,plazo) " _
        & "values('" & strCodigo & "'," & medDe.Text & "," & medHasta.Text & "," _
        & txtIntMorSocios.Text & "," & txtIntCorSocios.Text & "," _
        & txtIntMorNoSocios.Text & "," & txtIntCorNoSocios.Text & "," & txtPlazo.Text & ")"
   glogon.Conection.Execute strSQL
   Call Bitacora("Registra", "Rango para codigo" & strCodigo)
   If miTran = 1 Then
   glogon.Conection.CommitTrans
   miTran = 0
   End If
   'buscar columna
    With rs
     .Source = "select * from rangos where codigo = '" & strCodigo & "' and de = " & medDe.Text & " and hasta = " & medHasta.Text
     .CursorType = adOpenStatic
     .ActiveConnection = glogon.Conection
     .Open
     
     If .EOF = True And .BOF = True Then
      MsgBox "No se Guardó el Registro ...", vbCritical
      LimpiaPantalla
     Else
      Columna = !consec
     End If
     .Close
    End With
 Case 1 'edita
   strSQL = "update rangos set de = " & Val(medDe.Text) & "," _
          & "hasta = " & Val(medHasta.Text) & "," _
          & "intm_soc = " & txtIntMorSocios.Text & ", intc_soc = " & txtIntCorSocios.Text & "," _
          & "intm_nsoc = " & txtIntMorNoSocios.Text & ", intc_nsoc = " & txtIntCorNoSocios.Text & "," _
          & "plazo = " & txtPlazo.Text & " " _
          & "where consec = " & Columna
   glogon.Conection.Execute strSQL
   Call Bitacora("Modifica", "Rangos consec=" & Trim(Columna))
   Call ActualizaPrestamos
   'el commit lo hace el boton de actualiza
End Select

Else
  MsgBox "Los rangos incluidos no son válidos o ya existen ...", vbCritical
End If

Exit Sub
CapturaError:
   Me.MousePointer = vbDefault
   If miTran = 1 Then
   glogon.Conection.RollbackTrans
   miTran = 0
   End If
   Call ProcedimientoErrores(Me.Name, Err)
End Sub

Sub ActualizaPrestamos()
Dim i As Integer
'Compara la variable interes, para actualiza préstamos en reg_creditos
i = 1
intCambiaIntereses = 1

If Interes.MontoDe <> Val(medDe.Text) Then i = 0
If Interes.montohasta <> Val(medHasta.Text) Then i = 0

If i = 1 Then
 With Interes
  If .InteresCorrienteSocios <> Val(txtIntCorSocios.Text) Then strActualizaSocios = "S"
  If .InteresCorrienteNoSocios <> Val(txtIntCorNoSocios.Text) Then strActualizaSocios = "E"
 End With
Else
 'Actualiza todo
 strActualizaSocios = "T"
End If

tlbPrincipal.Enabled = False
Me.Height = Me.Height + fraActualizaIntereses.Height
fraActualizaIntereses.Visible = True

End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
Call Valida(KeyAscii)
End Sub
