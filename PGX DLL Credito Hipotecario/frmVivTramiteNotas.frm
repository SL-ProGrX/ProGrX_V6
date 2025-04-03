VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmVivTramiteNotas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notas de garantias"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread Grid 
      Height          =   3612
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   10932
      _Version        =   524288
      _ExtentX        =   19283
      _ExtentY        =   6371
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   486
      MaxRows         =   498
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmVivTramiteNotas.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Frame Frame4 
      Caption         =   "Información del crédito"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2892
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   11055
      Begin VB.TextBox txtNota 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1452
         Left            =   1800
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Label Label7 
         Caption         =   "N° Expediente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3840
         TabIndex        =   23
         Top             =   600
         Width           =   1092
      End
      Begin VB.Label lblExpediente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   5040
         TabIndex        =   22
         Top             =   576
         Width           =   1812
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle Nota"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmVivTramiteNotas.frx":066F
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Número Operación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label5 
         Caption         =   "Identificación:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   19
         Top             =   996
         Width           =   1452
      End
      Begin VB.Label Label1 
         Caption         =   "Información de la garantia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   18
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Área (m2)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7155
         TabIndex        =   17
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "N° plano catastro"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7155
         TabIndex        =   16
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "N° de finca"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7155
         TabIndex        =   15
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label lblPlano 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   3
         Left            =   7155
         TabIndex        =   13
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label lblArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   12
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label lblNumFinca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblDesZona 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblNumeroOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   1800
         TabIndex        =   9
         Top             =   576
         Width           =   1932
      End
      Begin VB.Label lblCedula 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   1932
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   3720
         TabIndex        =   7
         Top             =   960
         Width           =   3132
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   7155
         TabIndex        =   6
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label lblProvincia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Cantón"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   1
         Left            =   7155
         TabIndex        =   4
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label lblCanton 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8520
         TabIndex        =   3
         Top             =   2040
         Width           =   2175
      End
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVivTramiteNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vEditar As Boolean

Private m_cambioDatos As Boolean
Public m_NumOperacion As String
Public m_IdGarantia As Long
Public m_IdContacto As Long
Public m_Profesional As String
Public m_IdNota As Long

Private Sub sbLimpiarCampos()
txtNota.Text = Empty
End Sub
Private Sub sbTraerInformacionOperacion()
On Error GoTo vError
 
If ObjConsultar.fxTraerOperacionXIdGarantia(m_NumOperacion, m_IdGarantia) Then
lblNumeroOperacion.Caption = m_NumOperacion
 lblCedula.Caption = Trim(glogon.Recordset.Fields!cedula)
 lblNombre.Caption = (glogon.Recordset.Fields!Nombre)
 lblExpediente.Caption = IIf(IsNull(glogon.Recordset.Fields!Expediente), "", Trim(glogon.Recordset.Fields!Expediente))
 lblNumFinca.Caption = Trim(glogon.Recordset.Fields!NumeroFinca)
 lblPlano.Caption = IIf(IsNull(glogon.Recordset.Fields!NumPlanoCatastro), "", Trim(glogon.Recordset.Fields!NumPlanoCatastro))
 lblDesZona.Caption = Trim(glogon.Recordset.Fields!DescZona)
 lblArea.Caption = Trim(glogon.Recordset.Fields!AreaFinca)
 lblProvincia.Caption = Trim(glogon.Recordset.Fields!PROVINCIA)
 lblCanton.Caption = Trim(glogon.Recordset.Fields!Canton)
End If

salir:
    Exit Sub
vError:
    Call ObjMensajes.deError("Ocurrió un error en visual basic al consultar la información según número de operación. Error " & Err.Description)
End Sub

Private Sub Form_Activate()
vModulo = 3 'Modulo de Credito
End Sub

Private Sub Form_Load()
Grid.AppearanceStyle = fxGridStyle
vModulo = 3 'Modulo de Credito
'Inicializa Barra
Call sbToolBarIconos(tlbPrincipal, False)
Call sbToolBar(tlbPrincipal, "nuevo")
'Inicializa Seguridad
Call Formularios(Me)
Call RefrescaTags(Me)


Call sbTraerInformacionOperacion
Call sbRefrecaGrid(m_IdGarantia, m_Profesional)

End Sub



Private Sub Grid_DblClick(ByVal Col As Long, ByVal Row As Long)

    Grid.Row = Row
    Grid.Col = 2
    txtNota.Text = Grid.Text
    Grid.Col = 1
    m_IdNota = Grid.Text
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

    Select Case LCase(Button.Key)
            
        Case "nuevo"
            vEditar = False
            Call sbToolBar(Me.tlbPrincipal, "edicion")
            Call sbLimpiarCampos
        Case "editar"
            vEditar = True
            Call sbToolBar(Me.tlbPrincipal, "edicion")

        Case "borrar"
'            If (Grid.ActiveRow <> Grid.MaxRows) Then
'                Call sbBorrar(Grid.ActiveRow)
'            End If
        Case "guardar"
            Call sbGuardar
            
        Case "deshacer"
            vEditar = False
            Call sbToolBar(Me.tlbPrincipal, "nuevo")
    End Select
         

Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub
'Private Sub sbBorrar(ByRef pfila As Integer)
'On Error GoTo vError
'Dim vIdNota As Long
'Dim vIdGarantia As Long
'Dim vIdcontacto As Integer
'Dim vProfesional As String
'
'If m_ventanaEnModo = eVentanaEnModo.ModificarRegistro Then
'    Grid.Row = pfila
'    Grid.Col = 1
'    vIdNota = Grid.Text
'    Grid.Col = 6
'    vIdGarantia = Grid.Text
'    Grid.Col = 7
'    vIdcontacto = Grid.Text
'    Grid.Col = 8
'    vProfesional = Grid.Text
'
'    If ObjBorrar.fxNotasGarantiaTramite(vIdNota, vIdGarantia, vIdcontacto, vProfesional) Then
'        Call ObjMensajes.deDatos("06")
'        Grid.DeleteRows Grid.Row, 1
'        Grid.MaxRows = Grid.MaxRows - 1
'        Grid.Col = 1
'        Grid.SetActiveCell 1, Grid.ActiveRow
'        Call sbRefrecaGrid(m_IdGarantia, m_Profesional)
'        Call sbLimpiarCampos
'        Call SbAccionVentana(NuevoRegistro)
'
'    End If
'End If
'
'salir:
'    Exit Sub
'vError:
'    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
'    Resume salir
'End Sub

Private Sub sbGuardar()

If m_cambioDatos = False Then Exit Sub
   Call sbAgregar
'    Call sbModificar





End Sub

Private Sub sbRefrecaGrid(ByVal pIdGarantia As Long, ByVal pProfesional As String)

On Error GoTo vError

Me.MousePointer = vbHourglass
      
Grid.MaxCols = 10
Grid.MaxRows = 0

Grid.ColWidth(6) = 0
Grid.ColWidth(7) = 0
Grid.ColWidth(8) = 0
Grid.ColWidth(9) = 0
Grid.ColWidth(10) = 0

If ObjConsultar.fxNotasGarantiaTramite(pIdGarantia, pProfesional) Then
    Call sbCargaGrid(Grid, 10, glogon.strSQL)
End If

Grid.RowHeight(Grid.Row) = Grid.MaxTextRowHeight(Grid.Row)

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
Private Sub sbAgregar()
On Error GoTo vError
Me.MousePointer = vbHourglass
If fxValidaDatos = False Then Exit Sub
gParametros(0) = -1

If ObjAgregar.fxNotasGarantiaTramite(gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), gParametros(5)) Then
    m_cambioDatos = False
    MsgBox "Información fue registrada corretamente.", vbInformation
    Call sbRefrecaGrid(gParametros(1), gParametros(3))
    Call sbLimpiarCampos
    Call sbToolBar(tlbPrincipal, "Nuevo")
End If

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbModificar()
On Error GoTo error

Me.MousePointer = vbHourglass
If fxValidaDatos = False Then Exit Sub
gParametros(0) = m_IdNota
If ObjAgregar.fxNotasGarantiaTramite(gParametros(0), gParametros(1), gParametros(2), gParametros(3), gParametros(4), gParametros(5)) Then
    m_cambioDatos = False
    MsgBox "Información fue actualizada correctamente.", vbInformation
    Call sbRefrecaGrid(gParametros(1), gParametros(3))
    Call sbLimpiarCampos
    Call sbToolBar(tlbPrincipal, "Nuevo")

 End If
salir:
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Call ObjMensajes.deError("Ocurrió un error en visual basic al modificar la información ingresada. Error " & Err.Description)
End Sub
Private Function fxValidaDatos() As Boolean
On Error GoTo error

fxValidaDatos = False

ReDim gParametros(0 To 5)

If (Len(Trim(txtNota.Text)) = 0) Then
    Me.MousePointer = vbDefault
    MsgBox ("Debe de ingresar una nota válida.")
    txtNota.SetFocus
    Exit Function

End If

gParametros(1) = m_IdGarantia
gParametros(2) = m_IdContacto
gParametros(3) = m_Profesional
gParametros(4) = Trim(txtNota.Text)
gParametros(5) = glogon.Usuario



fxValidaDatos = True
salir:
    Exit Function
error:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub txtNota_Change()
m_cambioDatos = True
End Sub
