VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmPublishers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editores"
   ClientHeight    =   3090
   ClientLeft      =   765
   ClientTop       =   1230
   ClientWidth     =   7350
   Icon            =   "frmpub.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbout 
      Caption         =   "A&cerca de"
      Height          =   375
      Left            =   6180
      TabIndex        =   9
      Top             =   630
      Width           =   945
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6180
      TabIndex        =   8
      Top             =   240
      Width           =   945
   End
   Begin VB.OptionButton optListBox 
      Caption         =   "Usar un control ListBox estándar"
      Height          =   255
      Left            =   3660
      TabIndex        =   7
      Top             =   1860
      Width           =   3315
   End
   Begin VB.OptionButton optComboBox 
      Caption         =   "Usar un control ComboBox estándar"
      Height          =   255
      Left            =   3660
      TabIndex        =   6
      Top             =   1560
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.TextBox txtCity 
      DataSource      =   "datPublishers"
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Width           =   4000
   End
   Begin VB.TextBox txtAddress 
      DataSource      =   "datPublishers"
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   4000
   End
   Begin VB.TextBox txtCompanyName 
      DataSource      =   "datPublishers"
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   540
      Width           =   4000
   End
   Begin VB.TextBox txtName 
      DataSource      =   "datPublishers"
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4000
   End
   Begin VB.Data datPublishers 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   ""
      Top             =   2745
      Width           =   7344
   End
   Begin VB.ComboBox cboState 
      DataSource      =   "datPublishers"
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   4
      Text            =   "cmbState"
      Top             =   1455
      Width           =   1215
   End
   Begin VB.ListBox lstState 
      DataSource      =   "datPublishers"
      Height          =   675
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1455
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   420
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
     FontSize        =   1.17491e-38
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado:"
      Height          =   195
      Left            =   660
      TabIndex        =   14
      Top             =   1500
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Ciudad:"
      Height          =   195
      Left            =   660
      TabIndex        =   13
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Dirección:"
      Height          =   195
      Left            =   660
      TabIndex        =   12
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre de la compañía:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre:"
      Height          =   195
      Left            =   660
      TabIndex        =   10
      Top             =   300
      Width           =   1305
   End
End
Attribute VB_Name = "frmPublishers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
   Dim frmX As New frmAbout
   '
   ' Presenta el cuadro de diálogo Acerca de
   '
   frmX.Show vbModal
   Set frmX = Nothing
   
End Sub

Private Sub cmdDone_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim vntTemp As Variant
   Dim vnDummy As Variant
   
   On Error GoTo Form_LoadError
   
   ' Obtiene todos los estados de la tabla
   '
   datPublishers.DatabaseName = "biblio.mdb"
   datPublishers.RecordSource = "SELECT DISTINCT state FROM Publishers"
   datPublishers.Refresh
   '
   ' Carga todas las abreviaturas de estado en los controles cuadro de lista y cuadro combinado
   ' Si el nombre del archivo de la base de datos no es válido, aquí se produciría un error interceptable
   Do While Not datPublishers.Recordset.EOF
      vntTemp = datPublishers.Recordset!State
      If IsNull(vntTemp) Then vntTemp = ""
      cboState.AddItem CStr(vntTemp)
      lstState.AddItem CStr(vntTemp)
      datPublishers.Recordset.MoveNext
   Loop
   '
   ' Obtiene todos los registros de la tabla Publishers
   '
   datPublishers.RecordSource = "Publishers"
   datPublishers.RecordsetType = vbRSTypeTable
   datPublishers.Refresh
   '
   ' Enlaza los campos con el control Data
   '
   txtName.DataField = "Name"
   txtAddress.DataField = "Address"
   txtCompanyName.DataField = "Company Name"
   txtCity.DataField = "City"
   
   cboState.DataField = "State"
   lstState.DataField = "State"
    
Form_LoadExit:
Exit Sub
    
Form_LoadError:
    ' Si no se puede encontrar la base de datos Biblio, abre el control CommonDialog
    ' para que el usuario la busque. El archivo biblio.mdb se encuentra en el mismo
    ' directorio en el que vb5 ha sido instalado.
    If Err = 3024 Then
        With dlgDialog
            .DialogTitle = "Imposible encontrar Biblio.mdb"
            .Filter = "(*.MDB)|*.mdb"
            .ShowOpen
        End With
        
        If dlgDialog.filename <> "" Then
        'asegura que el archivo de base de datos devuelto es realmente biblio.mdb
            If Right(UCase(dlgDialog.filename), Len("\biblio.mdb")) = "\BIBLIO.MDB" Then
                datPublishers.DatabaseName = dlgDialog.filename
                datPublishers.RecordSource = "SELECT DISTINCT state FROM Publishers"
                datPublishers.Refresh
            End If
            Resume Next
        Else
            Unload Me
        End If
    ElseIf Err <> 0 Then ' otro error
        MsgBox "Error inesperado: " & Err.Description
        End
    End If
    Resume Form_LoadExit
   
End Sub


Private Sub optComboBox_Click()
   '
   ' alterna la visibilidad de los controles
   '
   cboState.Visible = True
   lstState.Visible = False
   
End Sub

Private Sub optListBox_Click()
   '
   ' alterna la visibilidad de los controles
   '
   cboState.Visible = False
   lstState.Visible = True

End Sub


