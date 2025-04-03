VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSIFCargaCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIF : Auxiliar para la creación de Créditos x Plantilla"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmSIFCargaCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSIFCargaCreditos.frx":6852
   ScaleHeight     =   4950
   ScaleWidth      =   6855
   Begin MSComctlLib.Toolbar tlbBuscar 
      Height          =   570
      Left            =   6000
      TabIndex        =   6
      Top             =   1680
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Archivo"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFCargaCreditos.frx":D0A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIFCargaCreditos.frx":22216
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox txtArchivo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Data DaoControl 
      Caption         =   "DaoControl"
      Connect         =   "dBASE IV;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2700
   End
   Begin MSComctlLib.Toolbar tlbProcesar 
      Height          =   780
      Left            =   5760
      TabIndex        =   7
      Top             =   4080
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1376
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Procesar"
            Key             =   "Procesar"
            Object.ToolTipText     =   "Procesar Archivo"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Picture         =   "frmSIFCargaCreditos.frx":37388
      Stretch         =   -1  'True
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmSIFCargaCreditos.frx":4C4EA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notas"
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
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   6840
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Carga de Créditos de Archivo Dbase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label lbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   6720
      X2              =   0
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmSIFCargaCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Function fxAnterior(xFecha As Long) As Long
Dim iMes As Integer, iAnio As Integer
Dim vFecha As String

vFecha = Trim(CStr(xFecha))

iAnio = Mid(vFecha, 1, 4)
iMes = Mid(vFecha, 5, 2)

If iMes = 1 Then
    iMes = 12
    iAnio = iAnio - 1
Else
    iMes = iMes - 1
End If

fxAnterior = iAnio & Format(iMes, "00")

End Function


Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim itmX As ListItem, vArchivo As String
Dim vPasa As Boolean


With frmContenedor.Dlg
 .InitDir = "C:\"
 .ShowOpen
 
 If .FileName = "" Then
   MsgBox "Archivo no válido...", vbExclamation
   Exit Sub
 End If
 
 If UCase(Right(.FileName, 3)) <> "DBF" Then
   MsgBox "La Extensión del Archivo no es válido...", vbExclamation
   Exit Sub
 End If

 txtArchivo = .FileName

End With

DaoControl.RecordSource = Dir(txtArchivo, vbArchive)
DaoControl.DatabaseName = Mid(txtArchivo, 1, Len(txtArchivo) - (Len(DaoControl.RecordSource) + 1))
DaoControl.Refresh


End Sub

Private Sub tlbProcesar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaSistema As Date, vCodigo As String
Dim lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vFechaSistema = fxFechaServidor

lng = 1

With DaoControl.Recordset

Do While Not .EOF

  lbl.Caption = "Procesando Registro : " & lng & " de " & .RecordCount + 1
  lbl.Refresh
  
  strSQL = "select coalesce(count(*),0) as existe from socios where cedula = '" & Trim(!cedula) & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs!existe = 0 Then
            strSQL = "Insert Socios(Cedula,Nombre,EstadoActual,FechaIngreso,Fecha_Nac," _
                   & "Sexo,EstadoCivil,Provincia,Canton,Distrito,Direccion,Af_Email,Apto," _
                   & "Cod_sector,cod_profesion,id_promotor,EstadoLaboral,Ultimo_Estado,Cod_Institucion" _
                   & ",Cod_departamento,cod_seccion,boleta,cedulaR,af_npagos,hijos,reg_user,reg_fecha) Values('" _
                   & Trim(!cedula) & "','" & Trim(!Nombre) & "'," & "'S','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
                   & Format(vFechaSistema, "yyyy/mm/dd") & "','M','S',1,'1','','','','',1,1,1,1,'N'," _
                   & !inst & ",'','',0,'',2,0,'" & glogon.Usuario & "',getdate())"
            glogon.Conection.Execute strSQL
           
            strSQL = "Insert into Ahorro_Consolidado(Cedula,Aporte,Ahorro,Extra,Capitaliza," _
                   & "FecAporte,FecAhorro,FecExtra,FecCapitaliza,AportAnt,AhorroAnt) Values(" _
                   & "'" & Trim(!cedula) & "',0,0,0,0,getdate(),getdate(),getdate(),getdate(),0,0)"
            glogon.Conection.Execute strSQL
   End If
   rs.Close
   
   'Saca un codigo por Institucion para el credito
   strSQL = "Select codigo from catalogo where cod_institucion = " & !inst
   rs.Open strSQL, glogon.Conection, adOpenStatic
    vCodigo = rs!codigo
   rs.Close
  
   '1. Insertar la Formalización
   strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
          & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
          & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
          & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
          & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,cod_destino,cod_grupo) values('" _
          & vCodigo & "',1,'" & Trim(!cedula) & "'," & !monto & "," & !monto _
          & ",0," & !monto & ",0,0," & !monto & "," & (!monto / !plazo) & "," & !tasa & "," _
          & !tasa & "," & !plazo & ",'" _
          & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" _
          & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
          & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','" _
          & Format(vFechaSistema, "yyyy/mm/dd") & "','" & Format(vFechaSistema, "yyyy/mm/dd") & "','N'" _
          & ",'N','OT','',0,1,0,'" & txtNotas & "','A'," & !PriDed _
          & "," & fxAnterior(!PriDed) & ",'F','','')"
   glogon.Conection.Execute strSQL
  
  lng = lng + 1
  .MoveNext
Loop

End With

Me.MousePointer = vbDefault

lbl.Caption = ""

MsgBox "Proceso finalizado Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub
