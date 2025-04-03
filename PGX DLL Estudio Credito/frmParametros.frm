VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PreAnalisis - Parámetros Generales"
   ClientHeight    =   6396
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8748
   HelpContextID   =   6002
   Icon            =   "frmParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   8748
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTabParametros 
      Height          =   6255
      Left            =   45
      TabIndex        =   1
      Top             =   120
      Width           =   8655
      _ExtentX        =   15261
      _ExtentY        =   11028
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Membresias"
      TabPicture(0)   =   "frmParametros.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vGrid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboGarantia"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lsw"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdActualizaCodigos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdActualizaTabla"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optOrden(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optOrden(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Otros"
      TabPicture(1)   =   "frmParametros.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(0)"
      Tab(1).Control(1)=   "Label3(1)"
      Tab(1).Control(2)=   "Label3(2)"
      Tab(1).Control(3)=   "Label4(0)"
      Tab(1).Control(4)=   "Label4(1)"
      Tab(1).Control(5)=   "Line2"
      Tab(1).Control(6)=   "Label4(2)"
      Tab(1).Control(7)=   "Label5(0)"
      Tab(1).Control(8)=   "Label5(1)"
      Tab(1).Control(9)=   "Line1"
      Tab(1).Control(10)=   "txtMesesTranscurridos"
      Tab(1).Control(11)=   "txtPorcentajeAhorroFid"
      Tab(1).Control(12)=   "txtPlazoCancelo"
      Tab(1).Control(13)=   "cmdActualizaOtros"
      Tab(1).Control(14)=   "chkActivaSGT"
      Tab(1).Control(15)=   "txtAutUsers"
      Tab(1).Control(16)=   "txtAutNumeros"
      Tab(1).Control(17)=   "lswAuto"
      Tab(1).Control(18)=   "cmdAutorizar"
      Tab(1).ControlCount=   19
      Begin VB.CommandButton cmdAutorizar 
         Caption         =   "&Autorizar"
         Height          =   855
         Left            =   -72840
         Picture         =   "frmParametros.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4560
         Width           =   1095
      End
      Begin MSComctlLib.ListView lswAuto 
         Height          =   2655
         Left            =   -71520
         TabIndex        =   24
         Top             =   3360
         Width           =   5055
         _ExtentX        =   8911
         _ExtentY        =   4678
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Auto.Pendientes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ult.Auto"
            Object.Width           =   1834
         EndProperty
      End
      Begin VB.TextBox txtAutNumeros 
         Height          =   315
         Left            =   -73320
         MaxLength       =   4
         TabIndex        =   23
         Top             =   3960
         Width           =   1695
      End
      Begin VB.TextBox txtAutUsers 
         Height          =   315
         Left            =   -73320
         MaxLength       =   4
         TabIndex        =   21
         Top             =   3600
         Width           =   1695
      End
      Begin VB.OptionButton optOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5760
         Width           =   3255
      End
      Begin VB.OptionButton optOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5760
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chkActivaSGT 
         Alignment       =   1  'Right Justify
         Caption         =   "Activar validación en el Seguimiento de Trámites ?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   16
         Top             =   2520
         Width           =   4815
      End
      Begin VB.CommandButton cmdActualizaOtros 
         Caption         =   "&Guardar"
         Height          =   855
         Left            =   -67920
         Picture         =   "frmParametros.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtPlazoCancelo 
         Height          =   315
         Left            =   -71040
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtPorcentajeAhorroFid 
         Height          =   315
         Left            =   -71040
         MaxLength       =   4
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtMesesTranscurridos 
         Height          =   315
         Left            =   -71040
         MaxLength       =   4
         TabIndex        =   0
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdActualizaTabla 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualiza Tabla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5760
         Width           =   2535
      End
      Begin VB.CommandButton cmdActualizaCodigos 
         Caption         =   "Actualiza Códigos"
         Height          =   255
         Left            =   6480
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4815
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4575
         _ExtentX        =   8065
         _ExtentY        =   8488
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.ComboBox cboGarantia 
         Height          =   315
         ItemData        =   "frmParametros.frx":0956
         Left            =   1200
         List            =   "frmParametros.frx":0966
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   3495
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4575
         Left            =   4800
         TabIndex        =   26
         Top             =   1080
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   8070
         _StockProps     =   64
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmParametros.frx":0998
         AppearanceStyle =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   -71640
         X2              =   -74880
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label5 
         Caption         =   "# Autorizaciones"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   22
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   20
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Autorización de Usuarios a Formalizar sobre Monto de Tabla de Membresias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   19
         Top             =   3120
         Width           =   7935
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74880
         X2              =   -66840
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label4 
         Caption         =   "Con Base a la Tabla de Membresía"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Con Base al Procentaje de Ahorros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   11
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje del Plazo Cancelado Crédito Anterior"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   10
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Meses Transcurridos Crédito Anterior "
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje para Fiduciarios (Base Ahorros)"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tabla de Montos x Membresía"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4800
         TabIndex        =   5
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Garantias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Function fxCheckCodigo(vCodigo As String, Vgar As String) As Boolean
Dim rs As New ADODB.Recordset

strSQL = "Select * From Pra_Codigos Where Codigo='" & vCodigo & "' "
strSQL = strSQL & "And Garantia='" & Vgar & "'"
With rs
  .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       fxCheckCodigo = True
    Else
       fxCheckCodigo = False
    End If
  .Close
End With

End Function

Function fxVerificaLinea(i As Integer)
vGrid.Row = i

vGrid.Col = 1
If vGrid.Text = "" Then
   fxVerificaLinea = False
Else
  vGrid.Col = 2
  If vGrid.Text = "" Then
     fxVerificaLinea = False
  Else
    vGrid.Col = 3
    If vGrid.Text = "" Then
       fxVerificaLinea = False
    Else
       fxVerificaLinea = True
    End If
  End If
End If


End Function

Private Sub sbCodigos()
Dim rs As New ADODB.Recordset, strGarantia As String
Dim itmX As ListItem, vGarantia As String

Me.MousePointer = vbHourglass
lsw.ListItems.Clear
vGrid.MaxRows = 0

vGarantia = ""

Select Case Trim(cboGarantia)
   Case "Fiduciaria"
     Call sbMembresia("F")
     strGarantia = "Where GAR_FIADORES='S'"
     vGarantia = "F"
   Case "Vivienda"
     Call sbMembresia("V")
     strGarantia = "Where GAR_HIPOTECA='S'"
     vGarantia = "V"
   Case "Especial"
     Call sbMembresia("E")
     strGarantia = "Where GAR_NO='N'"
     vGarantia = "E"

   Case "Sin Garantía"
     Call sbMembresia("S")
     strGarantia = "Where GAR_NO='S'"
     vGarantia = "S"

End Select

If vGarantia = "" Then
  Me.MousePointer = vbDefault
  Exit Sub
End If

strSQL = "Select Codigo,Descripcion From Catalogo " & strGarantia & " and (Retencion = 'N' and Poliza = 'N')"

Select Case True
  Case optOrden.Item(0).Value
     strSQL = strSQL & " order by codigo"
  Case optOrden.Item(1).Value
     strSQL = strSQL & " order by descripcion"
End Select

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , Trim(rs!Codigo))
      itmX.SubItems(1) = Trim(rs!Descripcion)
      itmX.Checked = fxCheckCodigo(Trim(rs!Codigo), vGarantia)
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub


Sub sbMembresia(vGarantia As String)
Dim rs As New ADODB.Recordset
Dim i As Integer

strSQL = "Select * From Pra_Membresias where Garantia='" & vGarantia & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
   vGrid.MaxRows = IIf(.RecordCount = 0, 1, .RecordCount)
   For i = 1 To .RecordCount
      vGrid.Row = i
      vGrid.Col = 1
      vGrid.Text = !Desde
      vGrid.Col = 2
      vGrid.Text = !Hasta
      vGrid.Col = 3
      vGrid.Text = Format(!Monto, "#############.00")
      
      .MoveNext
   Next i
 .Close
End With

End Sub


Private Sub cboGarantia_Click()
 Call sbCodigos
End Sub


Private Sub cmdActualizaOtros_Click()
Dim rs As New ADODB.Recordset

If Trim(txtMesesTranscurridos) = "" Or txtPorcentajeAhorroFid = "" Or txtPlazoCancelo = "" Then
 MsgBox "Faltan Datos", vbExclamation
Else
 Me.MousePointer = vbHourglass

 strSQL = "Select * From pra_Parametros"
 With rs
  .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = True Then
       strSQL = "Insert into Pra_Parametros(Meses_Transcurridos,Porc_Fiduciarios,Porc_Cancelado,ACTIVAR_SGT) values(" _
             & txtMesesTranscurridos & "," & txtPorcentajeAhorroFid & "," & txtPlazoCancelo & ",0)"
       Call ConectionExecute(strSQL)
       
    Else
       strSQL = "Update Pra_Parametros set Meses_Transcurridos = " & txtMesesTranscurridos _
              & ",Porc_Fiduciarios = " & txtPorcentajeAhorroFid _
              & ",Porc_Cancelado = " & txtPlazoCancelo & ",ACTIVAR_SGT = " & chkActivaSGT.Value
       Call ConectionExecute(strSQL)
    
    End If
  .Close
 End With
 
 Call Bitacora("Modifica", "Modifico Parametros De PreAnalisis")
 
 Me.MousePointer = vbDefault
 MsgBox "Registro Actualizado", vbExclamation
End If

End Sub

Private Sub cmdActualizaTabla_Click()
Dim i As Integer, strGarantia As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Trim(cboGarantia)
  Case "Fiduciaria"
    strGarantia = "'F'"
  Case "Vivienda"
    strGarantia = "'V'"
  Case "Especial"
    strGarantia = "'E'"
  Case "Sin Garantía"
    strGarantia = "'S'"
End Select

strSQL = "Delete Pra_Membresias where Garantia = " & strGarantia
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
 If fxVerificaLinea(i) = True Then
  vGrid.Row = i
  vGrid.Col = 1
  strSQL = "Insert Pra_Membresias(Garantia,Desde,Hasta,Monto) values(" & strGarantia & "," & vGrid.Text & ","
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 3
  strSQL = strSQL & Format(vGrid.Text, "###############.00") & ")"
  Call ConectionExecute(strSQL)
 End If
Next i

Call Bitacora("Modifica", "Modifico Tabla de Membresia Bajo Garantia " & Trim(cboGarantia))

Me.MousePointer = vbDefault

MsgBox "Tabla Actualizada", vbExclamation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub cmdAutorizar_Click()
Dim rs As New ADODB.Recordset

On Error GoTo vError


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset

vModulo = 3
Call Formularios(Me)

strSQL = "Select * from Pra_Parametros"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
      txtMesesTranscurridos = !Meses_Transcurridos
      txtPorcentajeAhorroFid = !Porc_Fiduciarios
      txtPlazoCancelo = !Porc_Cancelado
      chkActivaSGT.Value = IIf(IsNull(!ACTIVAR_SGT), 0, !ACTIVAR_SGT)
   End If
 .Close
End With

Call RefrescaTags(Me)
ssTabParametros.Tab = 0

lsw.Enabled = cmdActualizaCodigos.Enabled
vGrid.Enabled = cmdActualizaTabla.Enabled


End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim vGarantia As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Trim(cboGarantia.Text)
   Case "Fiduciaria"
     vGarantia = "F"
   Case "Vivienda"
     vGarantia = "V"
   Case "Especial"
     vGarantia = "E"
   Case "Sin Garantía"
     vGarantia = "S"
End Select

If Item.Checked Then
   strSQL = "insert pra_codigos(garantia,codigo) values('" & vGarantia & "','" & Item.Text & "')"
   Call Bitacora("Registra", "Registra Codigo " & Item.Text & " Bajo Garantia " & Trim(cboGarantia))
Else
   strSQL = "Delete From Pra_Codigos where Codigo='" & Item.Text & "' and Garantia = '" & vGarantia & "'"
   Call Bitacora("Borra", "Borra Codigo " & Item.Text & " Bajo Garantia " & Trim(cboGarantia))
End If
Call ConectionExecute(strSQL)


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub optOrden_Click(Index As Integer)
 Call sbCodigos
End Sub

Private Sub sbCargaAutorizados()
Dim strSQL As String, rs As New ADODB.Recordset






End Sub


Private Sub ssTabParametros_Click(PreviousTab As Integer)

If ssTabParametros.Tab = 1 Then
  Call sbCargaAutorizados
End If

End Sub

Private Sub txtMesesTranscurridos_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
     txtPorcentajeAhorroFid.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub


Private Sub txtPlazoCancelo_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
     cmdActualizaOtros.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub

Private Sub txtPorcentajeAhorroFid_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
     txtPlazoCancelo.SetFocus
  Case Else
     KeyAscii = 0
End Select
End Sub


Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
Dim intI As Integer

If vGrid.ActiveRow = 1 And vGrid.MaxRows > 1 Then
   Exit Sub
End If

Select Case vGrid.ActiveCol
  Case 3
    For intI = 1 To 3
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = intI

        If Trim(vGrid.Text) = "" Then
           Exit Sub
        End If
    Next
    
    vGrid.MaxRows = vGrid.MaxRows + 1
     Call gsbPulsarTecla(vbKeyTab)

End Select


End Sub


