VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_ConsultaMov 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Movimientos de Personas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAF_ConsultaMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   10770
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4932
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10572
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   8700
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
      Item(0).Caption =   "Ingresos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lswIngresos"
      Item(1).Caption =   "Renuncias"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lswRenuncia"
      Item(2).Caption =   "Liquidaciones"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "lswLiq"
      Item(2).Control(1)=   "cmdLiqReporte"
      Item(2).Control(2)=   "cmdLiqReversion"
      Item(2).Control(3)=   "lblLiq"
      Begin XtremeSuiteControls.ListView lswLiq 
         Height          =   3372
         Left            =   -69880
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   5948
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRenuncia 
         Height          =   4332
         Left            =   -69880
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   7641
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswIngresos 
         Height          =   4332
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   7641
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdLiqReporte 
         Height          =   612
         Left            =   -61480
         TabIndex        =   6
         Top             =   4128
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_ConsultaMov.frx":000C
      End
      Begin XtremeSuiteControls.PushButton cmdLiqReversion 
         Height          =   612
         Left            =   -63040
         TabIndex        =   7
         Top             =   4128
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reversión"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_ConsultaMov.frx":07C8
      End
      Begin VB.Label lblLiq 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -68320
         TabIndex        =   8
         Top             =   4080
         Visible         =   0   'False
         Width           =   2892
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Top             =   600
      Width           =   6495
      _Version        =   1441793
      _ExtentX        =   11456
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta de Movimientos de la Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1884
      TabIndex        =   1
      Top             =   120
      Width           =   6372
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   612
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmAF_ConsultaMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strEstado As String

Public Sub sbConsulaExterna(pCedula As String)

txtCedula.Text = pCedula
Call txtCedula_LostFocus

End Sub


Private Sub cmdLiqReporte_Click()

If lswLiq.Tag = "" Then Exit Sub

Call sbgAFIBoletaLiquidacion(lswLiq.Tag)

End Sub

Private Sub cmdLiqReversion_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

'Pasos:
'0. Verificar que la liquidacion sea la ultima activa
'1. Aplicar Reversion
'2. Refrescar Consulta


If Trim(lswLiq.Tag) = "" Then
   MsgBox "Especifique Un Número de Liquidación...", vbExclamation
   Exit Sub
End If


'Validacion: Nuevo Proceso desde el gestor
strSQL = "exec spAFI_Liquidacion_Reversa_Validacion " & lswLiq.Tag
Call OpenRecordSet(rs, strSQL)
If Len(rs!Mensaje) > 0 Then
   Me.MousePointer = vbDefault
   MsgBox rs!Mensaje, vbExclamation
   Exit Sub
End If
rs.Close

' Verificar que el aporte patronal o custodia para que sea igual al momento de la liquidacion
' de lo contrario no procede.

i = MsgBox("Está seguro que desea reversar esta liquidación ?", vbYesNo)
If i = vbNo Then
   Me.MousePointer = vbDefault
   Exit Sub
End If


'Reversa la Liquidacion
strSQL = "exec  spAFI_Liquidacion_Reversa " & lswLiq.Tag & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


Call Bitacora("Reversa", "Reversa Liquidación # " & lswLiq.Tag & " - Ced:" & txtCedula.Text)

'Actualiza Lista
Call sbLista_Carga(2)

Me.MousePointer = vbDefault
MsgBox "Reversión Finalizada Satisfactoriamente...", vbInformation

End Sub

Private Sub Form_Load()
vModulo = 1


Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

tcMain.Item(0).Selected = True

With lswIngresos.ColumnHeaders
  .Clear
  .Add , , "Fecha", 1440
  .Add , , "Boleta", 1440, vbCenter
  .Add , , "Promotor", 3600

End With


With lswRenuncia.ColumnHeaders
  .Clear
  .Add , , "Fecha", 1440
  .Add , , "Boleta", 1440, vbCenter
  .Add , , "Tipo", 1400, vbCenter
  .Add , , "Causa", 3500
End With

With lswLiq.ColumnHeaders
  .Clear
  .Add , , "Fecha", 1440
  .Add , , "Liq Id.", 1400
  .Add , , "Tipo", 1440, vbCenter
  .Add , , "Anterior", 1440, vbCenter
  .Add , , "Documento", 1440, vbCenter
  .Add , , "Ubicación", 1440, vbCenter
  .Add , , "Neto", 1640, vbRightJustify
  .Add , , "Estado", 1440, vbCenter

End With


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lswLiq_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vFecha As Date

On Error GoTo vError

vFecha = fxFechaServidor

lswLiq.Tag = Item.SubItems(1)
lblLiq.Caption = "Liquidación No. " & lswLiq.Tag

'Le permite reversar solo las procesadas en el mismo
'mes y año
If Month(CDate(lswLiq.SelectedItem.Text)) = Month(vFecha) _
   And Year(CDate(lswLiq.SelectedItem.Text)) = Year(vFecha) _
   And lswLiq.SelectedItem.SubItems(7) = "Procesada" Then

  cmdLiqReversion.Enabled = True
Else
  cmdLiqReversion.Enabled = False
End If

cmdLiqReporte.Enabled = True

Call RefrescaTags(Me)

Exit Sub

vError:
  lblLiq.Caption = ""
  lswLiq.Tag = ""
  cmdLiqReversion.Enabled = False
  cmdLiqReporte.Enabled = False


End Sub

Private Sub sbLista_Carga(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
  Case 0 'Reingresos
   
   tcMain.Item(0).Selected = True
  
   With lswIngresos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "ID", 900
        .ColumnHeaders.Add 2, , "Usuario", 1500
        .ColumnHeaders.Add 3, , "Fecha", 1500
        .ColumnHeaders.Add 4, , "Ingreso", 1500
        .ColumnHeaders.Add 5, , "Boleta", 1100
        .ColumnHeaders.Add 6, , "Promotor", 3500
     
        strSQL = "Select I.*,P.nombre as Promotor " _
               & " From Afi_Ingresos I left join promotores P on I.id_promotor = P.id_promotor" _
               & " where I.Cedula='" & Trim(txtCedula) & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!consec)
               itmX.SubItems(1) = rs!Usuario & ""
               itmX.SubItems(2) = rs!fecha & ""
               itmX.SubItems(3) = Format(rs!Fecha_Ingreso)
               itmX.SubItems(4) = rs!Boleta & ""
               itmX.SubItems(5) = rs!promotor & ""
           rs.MoveNext
        Loop
        rs.Close

   End With
  
 
  
  Case 1 'Renuncias
   tcMain.Item(1).Selected = True
   
   With lswRenuncia
     strSQL = "select R.*,C.descripcion" _
            & " from renuncias R left join causas_renuncias C on R.id_causa = C.id_causa" _
            & " where R.cedula = '" & txtCedula & "'"
     Call OpenRecordSet(rs, strSQL, 0)
     .ListItems.Clear
     
        
     Do While Not rs.EOF
     
       If GLOBALES.SysASEVersion Then
            If rs!TipoRen = "A" Then
             Set itmX = .ListItems.Add(, , Format(rs!FechaRena, "yyyy/mm/dd"))
                 itmX.SubItems(2) = "Ren.Asociación"
            Else
             Set itmX = .ListItems.Add(, , Format(rs!FechaRenP, "yyyy/mm/dd"))
                 itmX.SubItems(2) = "Ren.Patronal"
            End If
            
            itmX.SubItems(1) = rs!consec & ""
            itmX.SubItems(3) = rs!Descripcion & ""
     
       Else
            If rs!Tipo = "A" Then
             Set itmX = .ListItems.Add(, , Format(rs!fecha, "yyyy/mm/dd"))
                 itmX.SubItems(2) = "Ren.Asociación"
            Else
             Set itmX = .ListItems.Add(, , Format(rs!fecha, "yyyy/mm/dd"))
                 itmX.SubItems(2) = "Ren.Patronal"
            End If
            
            itmX.SubItems(1) = rs!consec & ""
            itmX.SubItems(3) = rs!Descripcion & ""
       End If
       rs.MoveNext
     Loop
     rs.Close
   End With
  
  
  Case 2 'Liquidaciones
   tcMain.Item(2).Selected = True
  
   cmdLiqReporte.Enabled = False
   'cmdLiqReversion.Enabled = False
   With lswLiq
     strSQL = "select L.*,E.descripcion as 'EstadoPersona'" _
            & " from liquidacion L inner join afi_estados_persona E on L.EstadoActual = E.cod_estado where L.cedula = '" & txtCedula & "' order by L.fecliq"
     Call OpenRecordSet(rs, strSQL, 0)
     .ListItems.Clear
     .Tag = ""
     lblLiq.Caption = ""
     Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , Format(rs!fecLiq, "yyyy/mm/dd"))
            itmX.SubItems(1) = rs!consec
       
       If rs!estadoactliq = "A" Then
            itmX.SubItems(2) = "Ren.Asociación"
       Else
            itmX.SubItems(2) = "Ren.Patronal"
       End If
                
       Select Case rs!EstadoActual
         Case "S"
            itmX.SubItems(3) = "Asociado"
         Case "A"
            itmX.SubItems(3) = "Ren.Asociación"
         Case "P"
            itmX.SubItems(3) = "Ren.Patronal"
         Case "N"
       End Select
       
       itmX.SubItems(3) = rs!EstadoPersona
       itmX.SubItems(4) = rs!TDOCUMENTO & rs!nDocumento & ""
       
       If rs!ubicacion = "C" Then
           itmX.SubItems(5) = "Contabilidad"
       Else
           itmX.SubItems(5) = "Tesorería"
       End If
       
       itmX.SubItems(6) = Format(rs!TNETO, "Standard")
       
       If rs!Estado = "P" Then
          itmX.SubItems(7) = "Procesada"
       Else
          itmX.SubItems(7) = "Reversada"
       End If
       
       rs.MoveNext
     Loop
     rs.Close
   End With
  
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call sbLista_Carga(Item.Index)
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Nombre"
   gBusquedas.Col3Name = "Id Alterno"

   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select Cedula,Nombre,CedulaR from socios"
   gBusquedas.Convertir = "N"
   
   frmBusquedas.Show vbModal
   
   txtCedula = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
   
   tcMain.Item(0).Selected = True
   Call sbLista_Carga(0)

End If

End Sub


Private Sub txtCedula_LostFocus()

txtNombre = fxNombre(txtCedula)
   
   tcMain.Item(0).Selected = True
   Call sbLista_Carga(0)

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Nombre"
   gBusquedas.Col3Name = "Id Alterno"
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Cedula,Nombre,Ced from socios"
   gBusquedas.Convertir = "N"
   
   frmBusquedas.Show vbModal
   
   txtCedula = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
   
   tcMain.Item(0).Selected = True
   Call sbLista_Carga(0)
   
End If

End Sub

Private Function fxInstitucion(strCedula As String) As Integer
Dim strSQL As String, rs As New ADODB.Recordset

'Codigo para extraer la institucion del asociado

strSQL = "select cod_institucion from socios where cedula  = '" & strCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Or Not rs.BOF Then
   fxInstitucion = rs!cod_institucion
Else
  fxInstitucion = 0
End If
rs.Close
End Function

