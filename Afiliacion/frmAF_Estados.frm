VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmAF_Estados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados de la Persona"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12045
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5772
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   12132
      _Version        =   1441793
      _ExtentX        =   21399
      _ExtentY        =   10181
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Estados"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Movimientos"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "cboEstadoFinal"
      Item(1).Control(1)=   "cboMovimiento"
      Item(1).Control(2)=   "cboEstadoInicial"
      Item(1).Control(3)=   "Label2(2)"
      Item(1).Control(4)=   "Label2(1)"
      Item(1).Control(5)=   "Label2(0)"
      Item(1).Control(6)=   "scMovimientos"
      Item(1).Control(7)=   "btnMovimientos(0)"
      Item(1).Control(8)=   "btnMovimientos(1)"
      Item(1).Control(9)=   "lsw"
      Item(2).Caption =   "Entidades"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "lswEstados"
      Item(2).Control(1)=   "ShortcutCaption1"
      Item(2).Control(2)=   "scEntidades"
      Item(2).Control(3)=   "lswEntidades"
      Item(2).Control(4)=   "chkEntidades"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3972
         Left            =   -69760
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
         Width           =   11532
         _Version        =   1441793
         _ExtentX        =   20341
         _ExtentY        =   7006
         _StockProps     =   77
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
      Begin XtremeSuiteControls.ListView lswEntidades 
         Height          =   4812
         Left            =   -65200
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   7092
         _Version        =   1441793
         _ExtentX        =   12509
         _ExtentY        =   8488
         _StockProps     =   77
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
      Begin XtremeSuiteControls.ListView lswEstados 
         Height          =   4812
         Left            =   -69880
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
         _ExtentY        =   8488
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkEntidades 
         Height          =   216
         Left            =   -65080
         TabIndex        =   17
         Top             =   560
         Visible         =   0   'False
         Width           =   216
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnMovimientos 
         Height          =   312
         Index           =   0
         Left            =   -60760
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Registrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmAF_Estados.frx":0000
         ImageAlignment  =   4
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5172
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   11292
         _Version        =   524288
         _ExtentX        =   19918
         _ExtentY        =   9123
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
         MaxCols         =   496
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_Estados.frx":0727
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoInicial 
         Height          =   312
         Left            =   -69760
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   3252
         _Version        =   1441793
         _ExtentX        =   5741
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboMovimiento 
         Height          =   312
         Left            =   -66520
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   5052
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoFinal 
         Height          =   312
         Left            =   -61480
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   3252
         _Version        =   1441793
         _ExtentX        =   5741
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnMovimientos 
         Height          =   312
         Index           =   1
         Left            =   -59560
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Eliminar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmAF_Estados.frx":0E6D
         ImageAlignment  =   4
      End
      Begin XtremeShortcutBar.ShortcutCaption scEntidades 
         Height          =   372
         Left            =   -65200
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   7092
         _Version        =   1441793
         _ExtentX        =   12509
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Entidades Asociadas a:  "
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   -69880
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Estados Disponibles"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption scMovimientos 
         Height          =   372
         Left            =   -69760
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   11532
         _Version        =   1441793
         _ExtentX        =   20341
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Movimientos Registrados:"
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
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Estado Inicial"
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
         Height          =   252
         Index           =   0
         Left            =   -69760
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   3252
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Movimiento"
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
         Height          =   252
         Index           =   1
         Left            =   -66520
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   5052
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Estado Final"
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
         Height          =   252
         Index           =   2
         Left            =   -61480
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   3252
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estados de la Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1884
      TabIndex        =   0
      Top             =   360
      Width           =   6132
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12252
   End
End
Attribute VB_Name = "frmAF_Estados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnMovimientos_Click(Index As Integer)
Dim strSQL As String, i As Integer
Dim pMovimiento As String


Me.MousePointer = vbHourglass

Select Case Index
  Case 0 'Guardar
     
     
     
     strSQL = "insert afi_estados_cambio(cod_estado,cod_movimiento,cod_estado_cambio,usuario,fecha) values('" _
            & cboEstadoInicial.ItemData(cboEstadoInicial.ListIndex) & "','" & cboMovimiento.ItemData(cboMovimiento.ListIndex) & "','" & cboEstadoFinal.ItemData(cboEstadoFinal.ListIndex) _
            & "','" & glogon.Usuario & "',dbo.MyGetdate())"
     Call ConectionExecute(strSQL)
     
     Call Bitacora("Registra", "Cambio Estado M." & cboMovimiento.ItemData(cboMovimiento.ListIndex) & " Ei." & cboEstadoInicial.ItemData(cboEstadoInicial.ListIndex) _
                & " Ef." & cboEstadoFinal.ItemData(cboEstadoFinal.ListIndex))
     
  Case 1 'Eliminar
    
    With lsw.ListItems
        For i = 1 To .Count
           If .Item(i).Checked Then
              
              
            Select Case Trim(.Item(i).SubItems(1))
              Case "Ingreso"
                pMovimiento = "ING"
              Case "Re-Ingreso"
                pMovimiento = "REI"
              Case "Renuncia"
                pMovimiento = "REN"
              Case "Liquidación"
                pMovimiento = "LIQ"
              Case "Activación"
                pMovimiento = "ACT"
            End Select
     
              strSQL = "delete afi_estados_cambio where cod_estado = '" & .Item(i).Tag & "' and cod_estado_cambio = '" _
                     & .Item(i).ToolTipText & "' and cod_movimiento = '" & pMovimiento & "'"
              Call ConectionExecute(strSQL)
              
              Call Bitacora("Elimina", "Cambio Estado M." & Mid(.Item(i).SubItems(1), 1, 3) & " Ei." & .Item(i).Tag & " Ef." & .Item(i).ToolTipText)
              
           End If
        Next i
    End With

End Select

Call sbCargaLswEstadosMov

Me.MousePointer = vbDefault


End Sub


Private Sub chkEntidades_Click()
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub
If scEntidades.Tag = "" Then Exit Sub
If lswEntidades.ListItems.Count <= 0 Then Exit Sub

If chkEntidades.Value = vbChecked Then
   strSQL = "insert into AFI_ESTADOS_INSTITUCIONES(cod_estado,cod_institucion,usuario,fecha)" _
          & " (select '" & scEntidades.Tag & "',cod_institucion,'" & glogon.Usuario & "',dbo.MyGetdate()" _
          & " from instituciones where activa = 1 and cod_institucion not in(select cod_institucion from AFI_ESTADOS_INSTITUCIONES" _
          & " where cod_estado = '" & scEntidades.Tag & "'))"
Else
   strSQL = "delete AFI_ESTADOS_INSTITUCIONES where cod_estado = '" & scEntidades.Tag & "'"
End If

Call ConectionExecute(strSQL)

Call sbEntidades_Load


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Estado Inicial", 2100
    .Add , , "Movimiento", 2100, vbCenter
    .Add , , "Estado Final", 2100
    
    .Add , , "Reg. Usuario", 2100, vbCenter
    .Add , , "Reg. Fecha", 2100, vbCenter
End With


With lswEstados.ColumnHeaders
    .Clear
    .Add , , "Estado de la Persona", 4600
End With

With lswEntidades.ColumnHeaders
    .Clear
    .Add , , "Entidad Descripción", 4000
    .Add , , "Desc.Corta", 3000
End With



Call Formularios(Me)
Call RefrescaTags(Me)

lswEntidades.Enabled = vGrid.Enabled
chkEntidades.Enabled = lswEntidades.Enabled

tcMain.Item(0).Selected = True

strSQL = "select cod_estado,descripcion,activo,deduce_creditos,deduce_patrimonio,deduce_ahorros from afi_estados_persona" _
      & " order by cod_estado"
Call sbCargaGrid(vGrid, 6, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from afi_estados_persona " _
       & " where cod_estado = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into afi_estados_persona(cod_estado,descripcion,activo,deduce_creditos,deduce_patrimonio,deduce_ahorros,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 4
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 5
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 6
  strSQL = strSQL & vGrid.Value & ",dbo.mygetdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Estado de Persona : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update afi_estados_persona set descripcion = '" & vGrid.Text & "',activo = "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & ", deduce_creditos = "
 vGrid.col = 4
 strSQL = strSQL & vGrid.Value & ", deduce_patrimonio = "
 vGrid.col = 5
 strSQL = strSQL & vGrid.Value & ", deduce_ahorros = "
 vGrid.col = 6
 strSQL = strSQL & vGrid.Value & ", actualiza_fecha = dbo.mygetdate(), actualiza_usuario = '" _
        & glogon.Usuario & "' where cod_estado = '"
 
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Estado de Persona : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub sbCargaLswEstadosMov()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
       
       
lsw.ListItems.Clear

strSQL = "select C.*,I.descripcion as EstadoInicial,F.descripcion as EstadoFinal" _
        & " from afi_estados_cambio C inner join afi_estados_persona I on C.cod_estado = I.cod_estado" _
        & " inner join afi_estados_persona F on C.cod_estado_cambio = F.cod_estado"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!EstadoInicial)
     itmX.Tag = rs!cod_estado
     
     Select Case Trim(rs!COD_MOVIMIENTO)
       Case "ING"
         itmX.SubItems(1) = "Ingreso"
       Case "REI"
         itmX.SubItems(1) = "Re-Ingreso"
       Case "REN"
         itmX.SubItems(1) = "Renuncia"
       Case "LIQ"
         itmX.SubItems(1) = "Liquidación"
       Case "ACT"
         itmX.SubItems(1) = "Activación"
     End Select
     
     itmX.SubItems(2) = rs!EstadoFinal
     
     itmX.ToolTipText = rs!cod_estado_cambio
     
     itmX.SubItems(3) = rs!Usuario
     itmX.SubItems(4) = rs!fecha
     
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub sbEntidades_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError


vPaso = True
lswEntidades.ListItems.Clear

strSQL = "select Inst.COD_INSTITUCION,Inst.Descripcion, Inst.DESC_CORTA" _
       & ", case when isnull(Est.COD_INSTITUCION,0) = 0 then 0 else 1 end as 'Check'" _
       & " from INSTITUCIONES Inst left join AFI_ESTADOS_INSTITUCIONES Est on Inst.COD_INSTITUCION = Est.COD_INSTITUCION" _
       & " and Est.COD_ESTADO = '" & scEntidades.Tag & "'" _
       & " Where Inst.ACTIVA = 1"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswEntidades.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = rs!desc_corta & ""
     itmX.Tag = rs!cod_institucion
     itmX.Checked = rs!Check
 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswEntidades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert AFI_ESTADOS_INSTITUCIONES(cod_estado,cod_institucion,usuario,fecha) values('" & scEntidades.Tag _
          & "'," & Item.Tag & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = "delete AFI_ESTADOS_INSTITUCIONES where cod_estado = '" & scEntidades.Tag & "' and cod_institucion = " & Item.Tag
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub lswEstados_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
scEntidades.Caption = "Entidades Relacionadas a: " & Item.Text
scEntidades.Tag = Item.Tag

Call sbEntidades_Load
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

Me.MousePointer = vbHourglass

Select Case Item.Index

  Case 0

        strSQL = "select cod_estado,descripcion,activo,deduce_creditos,deduce_patrimonio,deduce_ahorros from afi_estados_persona" _
              & " order by cod_estado"
        Call sbCargaGrid(vGrid, 6, strSQL)
  
  Case 1 'Movimientos
        strSQL = "select rtrim(cod_estado) as 'IdX',  rtrim(descripcion) as 'ItmX' from afi_estados_persona"
        Call sbCbo_Llena_New(cboEstadoInicial, strSQL, False, True)
        Call sbCbo_Copia(cboEstadoInicial, cboEstadoFinal)
        With cboMovimiento
           .Clear
           .AddItem "Ingreso"
           .ItemData(.ListCount - 1) = "ING"
           
           .AddItem "Re-Ingreso"
           .ItemData(.ListCount - 1) = "REI"
           
           .AddItem "Renuncia"
           .ItemData(.ListCount - 1) = "REN"
           
           .AddItem "Liquidación"
           .ItemData(.ListCount - 1) = "LIQ"
           
           .AddItem "Activación"
           .ItemData(.ListCount - 1) = "ACT"
           
           .Text = "Ingreso"
        End With
        
        Call sbCargaLswEstadosMov
        
  Case 2 'Instituciones
     
     scEntidades.Tag = ""
     scEntidades.Caption = "Indique un Estado a Relacionar!"
     
     lswEntidades.ListItems.Clear
     lswEstados.ListItems.Clear
     
     strSQL = "select cod_estado,descripcion from afi_estados_persona where activo = 1"
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
      Set itmX = lswEstados.ListItems.Add(, , Trim(rs!Descripcion))
          itmX.Tag = rs!cod_estado
      rs.MoveNext
     Loop
     rs.Close
   
  
End Select

Me.MousePointer = vbDefault

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete afi_estados_persona where cod_estado = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Estado de Persona : " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


