VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_APA_Cortes 
   Caption         =   "Cortes por Acreedor"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cortes"
      TabPicture(0)   =   "frmCR_APA_Cortes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStatus"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblAcreedor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "prgGeneraCorte"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tlbCortes"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpFechaCorte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lsw"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkTodos"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Cierres"
      TabPicture(1)   =   "frmCR_APA_Cortes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tlbCierre"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "vGrid"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
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
         Left            =   -74880
         TabIndex        =   8
         Top             =   780
         Width           =   1455
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3705
         Left            =   -74880
         TabIndex        =   1
         Top             =   1140
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   6535
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3840
         Left            =   480
         TabIndex        =   3
         Top             =   540
         Width           =   7725
         _Version        =   524288
         _ExtentX        =   13626
         _ExtentY        =   6773
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   460
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_Cortes.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComCtl2.DTPicker dtpFechaCorte 
         Height          =   315
         Left            =   -70440
         TabIndex        =   4
         Top             =   1500
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   117702657
         CurrentDate     =   40559
      End
      Begin MSComctlLib.Toolbar tlbCierre 
         Height          =   330
         Left            =   6120
         TabIndex        =   7
         Top             =   4620
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         ButtonWidth     =   2672
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Realizar Cierre"
               Key             =   "Cerrar"
               Object.ToolTipText     =   "Agregar Acreedor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCortes 
         Height          =   330
         Left            =   -68880
         TabIndex        =   9
         Top             =   3720
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   582
         ButtonWidth     =   2646
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Generar Corte"
               Key             =   "Corte"
               Object.ToolTipText     =   "Agregar Acreedor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prgGeneraCorte 
         Height          =   255
         Left            =   -70800
         TabIndex        =   11
         Top             =   2760
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Acreedor:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   13
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lblAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -70560
         TabIndex        =   12
         Top             =   1920
         Width           =   3855
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -70560
         TabIndex        =   10
         Top             =   2280
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   6
         Top             =   2220
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   5
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   -74880
         X2              =   -72240
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   -72240
         X2              =   -72240
         Y1              =   420
         Y2              =   4980
      End
      Begin VB.Label Label1 
         Caption         =   "Acreedores"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   2
         Top             =   420
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCR_APA_Cortes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FechaServer As String, strSQL As String
Dim rs As New ADODB.Recordset
Dim mCod_Acreedor As String, mOperacion As String
Dim mNPago As Integer

Private Sub chkTodos_Click()
Dim I As Integer
 If chkTodos.Value = 1 Then
    For I = lsw.ListItems.Count To 1 Step -1
      lsw.ListItems(I).Checked = True
    Next I
 Else
    For I = lsw.ListItems.Count To 1 Step -1
      lsw.ListItems(I).Checked = False
    Next I
 End If
 
End Sub

Private Sub Form_Activate()
  vModulo = 3
End Sub

Private Sub Form_Load()
  vModulo = 3
  
  SSTab.Tab = 0
  
  FechaServer = fxFechaServidor
  dtpFechaCorte.Value = FechaServer
  vGrid.MaxCols = 5
  prgGeneraCorte.Value = 0
  prgGeneraCorte.Max = 2
  vGrid.MaxRows = 0
  
  Call sbCargaAcreedores
  
End Sub

Private Sub sbCargaAcreedores()
On Error GoTo vError
Dim vItem As MSComctlLib.ListItem, vLvw As MSComctlLib.ListView
Dim vKey As String

Me.lsw.ColumnHeaders.Clear
Me.lsw.ListItems.Clear

Set vLvw = Me.lsw
vLvw.ColumnHeaders.Add , , "Descripción", 2400
strSQL = "select COD_ACREEDOR, DESCRIPCION, ESTADO  from dbo.CRD_APA_ACREEDORES " & _
         " order by COD_ACREEDOR asc "
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
    vKey = Trim(rs.Fields("COD_ACREEDOR")) & "(CA)"
    Set vItem = lsw.ListItems.Add(, vKey, Trim(rs.Fields!DESCRIPCION))
    rs.MoveNext
Loop
rs.Close

Exit Sub
vError:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
On Error GoTo vError

mCod_Acreedor = Empty
mOperacion = Empty
Call sbListaMarcarSoloUno(lsw, Item)
If Item.Checked = True Then
   mCod_Acreedor = DeCodificaPrimaryKey(Item.Key, 1, "(CA)")
End If
Call sbCargarListaCortes

SSTab.Tab = 0
chkTodos.Value = 0
Exit Sub
    
vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub sbCargarListaCortes()
' Carga Lista de operaciones
Dim strSQL As String
Dim I As Integer, vGarantias As Currency, vResponsabilidad As Currency
Dim vFechaUltimoCorte As String

On Error GoTo vError

strSQL = "Select "

With vGrid
'Consulta la lista de las Operaciones
'Trae el último corte que este abierto
strSQL = "select O.COD_ACREEDOR,O.OPERACION,O.SALDO,GC.FECHA_CORTE as 'FechaCorte',isnull([dbo].[fxCRDAPASaldoResponsabilidad](O.COD_ACREEDOR, O.OPERACION),0) as 'Responsabilidad' " _
       & ",isnull([dbo].[fxCRDAPASaldoCorteGarantias] (O.COD_ACREEDOR ,O.OPERACION,max(GC.FECHA_CORTE)),0) as 'Garantias' " _
       & " from CRD_APA_OPERACIONES O inner join CRD_APA_GARANTIAS_CORTES GC on O.COD_ACREEDOR=GC.COD_ACREEDOR and O.OPERACION=GC.OPERACION" _
       & " where O.COD_ACREEDOR = '" & Trim(mCod_Acreedor) & "' and GC.ESTADO='A'" _
       & " group by O.COD_ACREEDOR,O.OPERACION,O.SALDO,GC.FECHA_CORTE"
       
rs.Open strSQL, glogon.Conection, adOpenStatic
    
.MaxRows = 1

Do While Not rs.EOF
  .Row = .MaxRows
  .Col = 2
  .Text = rs!Operacion
     
  .Col = 3
   vResponsabilidad = rs!Responsabilidad
  .Text = Format(vResponsabilidad, "Standard")
  
  .Col = 4
  vGarantias = rs!Garantias
  .Text = Format(vGarantias, "Standard")
  
  .Col = 5
  .Text = rs!FechaCorte
         
  If vGarantias >= vResponsabilidad Then
    .Col = 1
    .Value = 1
  End If
  
  .MaxRows = .MaxRows + 1
  rs.MoveNext
Loop
rs.Close
.MaxRows = .MaxRows - 1

End With

Exit Sub
vError:
  MsgBox Err.Description, vbCritical
    
End Sub

Private Sub sbAgregarCorte()
Dim strSQL As String, rs As New ADODB.Recordset, strSqltmp As String, mCod_Acreedor As String
Dim rsTmp As New ADODB.Recordset
Dim I As Integer

On Error GoTo vError


prgGeneraCorte.Value = 1
prgGeneraCorte.Max = 2


'Recorre la lista de acreedores y crea el corte de los q se encuentra marcados
For I = lsw.ListItems.Count To 1 Step -1
    If lsw.ListItems(I).Checked Then
       mCod_Acreedor = DeCodificaPrimaryKey(lsw.ListItems(I).Key, 1, "(CA)")
       strSQL = "Select COD_ACREEDOR,OPERACION from CRD_APA_OPERACIONES" _
              & " where COD_ACREEDOR='" & mCod_Acreedor & "' "
              
       rs.Open strSQL, glogon.Conection, adOpenStatic
       
       prgGeneraCorte.Max = rs.RecordCount + 1
    
       Do While Not rs.EOF
          'Crea cortes mientras no tenga ninguno abierto
            strSqltmp = "exec spCRDAPAGARANTIASCORTES_A " & pcc(mCod_Acreedor) _
                                                          & pcc(rs!Operacion) _
                                                          & pcc(Format(dtpFechaCorte.Value, "yyyymmdd")) _
                                                          & pcc(Format(FechaServer, "yyyymmdd hh:mm:ss")) _
                                                          & pcc(glogon.Usuario) _
                                                          & pc("Corte Automatico")
            glogon.Conection.Execute strSqltmp
            
            If prgGeneraCorte.Max > prgGeneraCorte.Value Then prgGeneraCorte.Value = prgGeneraCorte.Value + 1
            lblAcreedor.Caption = lsw.ListItems(I).Text
            lblStatus.Caption = "Generando Cortes..Registro # " & prgGeneraCorte.Value & " de " & prgGeneraCorte.Max & "     " & Format((prgGeneraCorte.Value / prgGeneraCorte.Max) * 100, "##0") & "%"
            lblStatus.Refresh
         
         rs.MoveNext
       Loop
       rs.Close
     End If 'Item.Checked = True
Next I

Me.MousePointer = vbDefault
MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox "Ocurrió un error en visual basic al agregar la información ingresada. Error " & Err.Description
    
End Sub

'Cierra el corte
Private Sub tlbCierre_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
Dim I As Integer, vOperacion As String, vFechaCorte As String
Me.MousePointer = vbHourglass
With vGrid
For I = 1 To .MaxRows
  .Col = 1
  .Row = I
  If .Value = 1 Then
    .Col = 2
    vOperacion = .Text
    .Col = 5
    vFechaCorte = Format(.Text, "yyyymmdd")
    strSQL = "exec spCRDAPAGARANTIASCORTES_CERRAR " & pcc(mCod_Acreedor) & _
                                                pcc(vOperacion) & _
                                                pcc(vFechaCorte) & _
                                                pcc(Format(FechaServer, "yyyymmdd hh:mm:ss")) & _
                                                pc(glogon.Usuario)
    rs.Open strSQL, glogon.Conection, adOpenStatic
  End If
Next I
End With
Call sbCargarListaCortes
MsgBox "Corte cerrado satisfactoriamente...", vbInformation
Me.MousePointer = vbDefault
Exit Sub
vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub tlbCortes_ButtonClick(ByVal Button As MSComctlLib.Button)
   Call sbAgregarCorte
End Sub
