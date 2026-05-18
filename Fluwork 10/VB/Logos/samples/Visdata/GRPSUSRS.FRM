VERSION 5.00
Begin VB.Form frmGroupsUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupos, usuarios y permisos"
   ClientHeight    =   5010
   ClientLeft      =   4380
   ClientTop       =   1905
   ClientWidth     =   6210
   HelpContextID   =   2016088
   Icon            =   "GRPSUSRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboOwners 
      Height          =   315
      Left            =   3780
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2360
      Width           =   2295
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Permisos"
      Height          =   1695
      Left            =   2820
      TabIndex        =   16
      Top             =   2760
      Width           =   3255
      Begin VB.CommandButton cmdAssign 
         Caption         =   "&Asignar"
         Height          =   300
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   25
         Top             =   1280
         Width           =   1400
      End
      Begin VB.CheckBox chkDeleteData 
         Caption         =   "Eliminar datos"
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkInsertData 
         Caption         =   "Insertar datos"
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   23
         Top             =   1000
         Width           =   1335
      End
      Begin VB.CheckBox chkUpdateData 
         Caption         =   "Actualizar datos"
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   22
         Top             =   660
         Width           =   1515
      End
      Begin VB.CheckBox chkReadData 
         Caption         =   "Leer datos"
         Height          =   255
         Left            =   1680
         MaskColor       =   &H00000000&
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox chkAdminister 
         Caption         =   "Administrar"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   20
         Top             =   1000
         Width           =   1455
      End
      Begin VB.CheckBox chkModifyDesign 
         Caption         =   "Modificar diseño"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   19
         Top             =   680
         Width           =   1575
      End
      Begin VB.CheckBox chkReadDesign 
         Caption         =   "Leer diseño"
         Height          =   255
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.OptionButton optGroups 
      Caption         =   "Grupos"
      Height          =   255
      Left            =   1200
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optUsers 
      Caption         =   "Usuarios"
      Height          =   255
      Left            =   120
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   2460
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox lstTablesQuerys 
      Height          =   1650
      ItemData        =   "GRPSUSRS.frx":030A
      Left            =   2820
      List            =   "GRPSUSRS.frx":030C
      MultiSelect     =   2  'Extended
      TabIndex        =   15
      Top             =   480
      Width           =   3255
   End
   Begin VB.PictureBox picUsers 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   2445
      TabIndex        =   3
      Top             =   480
      Width           =   2445
      Begin VB.CommandButton cmdPassword 
         Caption         =   "E&stablecer/Borrar contraseña"
         Height          =   300
         Left            =   15
         MaskColor       =   &H00000000&
         TabIndex        =   26
         Top             =   3615
         Width           =   2400
      End
      Begin VB.ListBox lstUsersGroups 
         Height          =   1065
         ItemData        =   "GRPSUSRS.frx":030E
         Left            =   0
         List            =   "GRPSUSRS.frx":0310
         MultiSelect     =   1  'Simple
         TabIndex        =   12
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Top             =   1800
         Width           =   1200
      End
      Begin VB.CommandButton cmdNewUser 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   9
         Top             =   1800
         Width           =   1200
      End
      Begin VB.ListBox lstUsers 
         Height          =   1650
         ItemData        =   "GRPSUSRS.frx":0312
         Left            =   0
         List            =   "GRPSUSRS.frx":0314
         TabIndex        =   4
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grupos pertenecen a:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.PictureBox picGroups 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   2325
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   2325
      Begin VB.ListBox lstGroupsUsers 
         Height          =   1065
         ItemData        =   "GRPSUSRS.frx":0316
         Left            =   0
         List            =   "GRPSUSRS.frx":0318
         MultiSelect     =   1  'Simple
         TabIndex        =   13
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton cmdDeleteGroup 
         Caption         =   "&Eliminar"
         Height          =   300
         Left            =   1200
         MaskColor       =   &H00000000&
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdNewGroup 
         Caption         =   "&Nuevo"
         Height          =   300
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox lstGroups 
         Height          =   1650
         ItemData        =   "GRPSUSRS.frx":031A
         Left            =   0
         List            =   "GRPSUSRS.frx":031C
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Miembros:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Propietario:"
      Height          =   255
      Index           =   3
      Left            =   2820
      TabIndex        =   27
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tablas y consultas:"
      Height          =   255
      Index           =   0
      Left            =   2820
      TabIndex        =   17
      Top             =   165
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2700
      X2              =   2700
      Y1              =   120
      Y2              =   4440
   End
End
Attribute VB_Name = "frmGroupsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'>>>>>>>>>>>>>>>>>>>>>>>>
Const FORMCAPTION = "Grupos/Usuarios/Permisos"
Const Label1 = "Tablas/Consultas:"
Const Label2 = "Grupos pertenecen a:"
Const LABEL3 = "Miembros:"
Const LABEL4 = "Propietario:"
Const BUTTON1 = "&Nuevo"
Const BUTTON2 = "&Eliminar"
Const BUTTON3 = "E&stablecer/Borrar contraseña"
Const BUTTON4 = "&Asignar"
Const BUTTON5 = "&Cerrar"
Const OPTION1 = "Usuarios"
Const OPTION2 = "Grupos"
Const FRAME1 = "Permisos"
Const MSG1 = "No tiene permiso para cambiar el propietario"
Const MSG2 = "No hay grupos seleccionados"
Const MSG3 = "¿Desea eliminar el grupo?"
Const MSG4 = "No hay usuarios seleccionados"
Const MSG5 = "¿Desea eliminar el usuario?"
Const MSG6 = "Nuevo grupo"
Const MSG7 = "Nuevo usuario"
Const MSG8 = "¿Desea borrar la contraseña?"
Const MSG9 = "No hay ningún objeto seleccionado"
'>>>>>>>>>>>>>>>>>>>>>>>>


Dim mbSettingUser As Integer
Dim mbSettingOwner As Integer
Dim mbSettingPerm As Integer
Dim mbLoading As Integer
Dim mobjCurrObject As Object    'tabla o consulta actualmente seleccionada

Private Sub cboOwners_Click()
  On Error GoTo COErr
  
  'si estamos estableciendo desde el código, sale
  If mbSettingOwner Then Exit Sub
  
  If (mobjCurrObject.Permissions And dbSecWriteOwner) = dbSecWriteOwner Then
    'intenta establecerlo
    mobjCurrObject.Owner = cboOwners.Text
  Else
    MsgBox MSG1, 48
    Exit Sub
  End If
  
  Exit Sub
  
COErr:
  ShowError
End Sub

Private Sub chkAdminister_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkAdminister.Value = vbChecked Then
    'establece todos
    chkReadDesign.Value = vbChecked
    chkModifyDesign.Value = vbChecked
    chkReadData.Value = vbChecked
    chkUpdateData.Value = vbChecked
    chkInsertData.Value = vbChecked
    chkDeleteData.Value = vbChecked
  End If
End Sub

Private Sub chkDeleteData_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkDeleteData.Value = vbUnchecked Then
    'desactiva los que los necesitan
    chkAdminister.Value = vbUnchecked
    chkModifyDesign.Value = vbUnchecked
  Else
    chkReadDesign.Value = vbChecked
    chkReadData.Value = vbChecked
  End If
End Sub

Private Sub chkInsertData_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkInsertData.Value = vbUnchecked Then
    ' desactiva los que los necesitan
    chkAdminister.Value = vbUnchecked
  Else
    chkReadDesign.Value = vbChecked
    chkReadData.Value = vbChecked
  End If
End Sub

Private Sub chkModifyDesign_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkModifyDesign.Value = vbUnchecked Then
    'desactiva su administración
    chkAdminister.Value = vbUnchecked
  Else
    chkReadDesign.Value = vbChecked
    chkReadData.Value = vbChecked
    chkInsertData.Value = vbChecked
    chkDeleteData.Value = vbChecked
  End If
End Sub

Private Sub chkReadData_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkReadData.Value = vbUnchecked Then
    'desactiva los que los necesitan
    chkAdminister.Value = vbUnchecked
    chkModifyDesign.Value = vbUnchecked
  Else
    chkReadDesign.Value = vbChecked
  End If
End Sub

Private Sub chkReadDesign_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkReadDesign.Value = vbUnchecked Then
    'desactiva todos
    chkAdminister.Value = vbUnchecked
    chkModifyDesign.Value = vbUnchecked
    chkReadData.Value = vbUnchecked
    chkUpdateData.Value = vbUnchecked
    chkInsertData.Value = vbUnchecked
    chkDeleteData.Value = vbUnchecked
  End If
End Sub

Private Sub chkUpdateData_Click()
  If mbSettingPerm Then Exit Sub
  
  If chkUpdateData.Value = vbUnchecked Then
    'desactiva los que los necesitan
    chkAdminister.Value = vbUnchecked
    chkModifyDesign.Value = vbUnchecked
  Else
    chkReadDesign.Value = vbChecked
    chkReadData.Value = vbChecked
  End If
End Sub

Private Sub cmdAssign_Click()
  SetPermissions True    'los asignará
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdDeleteGroup_Click()
  On Error GoTo DGErr
  
  Dim i As Integer
    
  If lstGroups.ListIndex < 0 Then
    Beep
    MsgBox MSG2
    Exit Sub
  End If
  If MsgBox(MSG3, vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  
  gwsMainWS.Groups.Delete lstGroups.Text
  i = lstGroups.ListIndex
  lstGroups.RemoveItem i
  lstUsersGroups.RemoveItem i
  If lstGroups.ListCount > 0 Then
    lstGroups.ListIndex = 0
  Else
    'necesario para deseleccionar todos los usuarios
    For i = 0 To lstGroupsUsers.ListCount - 1
      lstGroupsUsers.Selected(i) = False
    Next
  End If
  
  Exit Sub
  
DGErr:
  ShowError
End Sub

Private Sub cmdDeleteUser_Click()
  On Error GoTo DUErr
  
  Dim i As Integer
  
  If lstUsers.ListIndex < 0 Then
    Beep
    MsgBox MSG4
    Exit Sub
  End If
  If MsgBox(MSG5, vbYesNo + vbQuestion) <> vbYes Then Exit Sub
  
  gwsMainWS.Users.Delete lstUsers.Text
  lstUsers.RemoveItem lstUsers.ListIndex
  If lstUsers.ListCount > 0 Then
    lstUsers.ListIndex = 0
  Else
    'necesario para deseleccionar todos los grupos
    For i = 0 To lstUsersGroups.ListCount - 1
      lstUsersGroups.Selected(i) = False
    Next
  End If
  Exit Sub
  
DUErr:
  ShowError
End Sub

Private Sub cmdNewGroup_Click()
  frmNewUserGroup.UserOrGroup = 1
  frmNewUserGroup.Caption = MSG6
  frmNewUserGroup.Show vbModal
End Sub

Private Sub cmdNewUser_Click()
  frmNewUserGroup.UserOrGroup = 0
  frmNewUserGroup.Caption = MSG7
  frmNewUserGroup.Show vbModal
End Sub

Private Sub cmdPassword_Click()
  On Error GoTo CPErr
  
  If lstUsers.Text = gwsMainWS.UserName Then
    frmNewPassword.Show vbModal
  Else
    If MsgBox(MSG8, vbYesNo + vbQuestion) = vbYes Then
      gwsMainWS.Users(lstUsers.Text).NewPassword vbNullString, vbNullString
    End If
  End If
  
  Exit Sub
  
CPErr:
  ShowError
End Sub

Private Sub Form_Load()
  On Error GoTo FLErr
  
  Dim grp As Group
  Dim usr As User
  Dim i As Integer

  Me.Caption = FORMCAPTION
  optUsers.Caption = OPTION1
  optGroups.Caption = OPTION2
  fraPermissions.Caption = FRAME1
  cmdNewUser.Caption = BUTTON1
  cmdDeleteUser.Caption = BUTTON2
  cmdNewGroup.Caption = BUTTON1
  cmdDeleteGroup.Caption = BUTTON2
  cmdPassword.Caption = BUTTON3
  cmdAssign.Caption = BUTTON4
  cmdClose.Caption = BUTTON5
  lblLabels(0).Caption = Label1
  lblLabels(1).Caption = Label2
  lblLabels(2).Caption = LABEL3
  lblLabels(3).Caption = LABEL4

  mbLoading = True
  'agrega los grupos y los usuarios
  For Each usr In gwsMainWS.Users
    lstUsers.AddItem usr.Name
    lstGroupsUsers.AddItem usr.Name
    cboOwners.AddItem usr.Name
  Next
  For Each grp In gwsMainWS.Groups
    lstGroups.AddItem grp.Name
    lstUsersGroups.AddItem grp.Name
    cboOwners.AddItem grp.Name
  Next
  'establece el primer elemento si es posible
  If lstUsers.ListCount > 0 Then
    lstUsers.ListIndex = 0
  End If
  If lstGroups.ListCount > 0 Then
    lstGroups.ListIndex = 0
  End If
  
  'llena las listas de objetos desde el formulario de las tablas
  GetTableList lstTablesQuerys, True, False, False
  mbLoading = False
  lstTablesQuerys.Selected(0) = True
  Screen.MousePointer = vbDefault
  Exit Sub
  
FLErr:
  mbLoading = False
  ShowError
End Sub

Private Sub lstGroups_Click()
  On Error GoTo GSErr

  Dim i As Integer
  
  mbSettingUser = True
  For i = 0 To lstGroupsUsers.ListCount - 1
    If IsMemberOf(lstGroups.Text, lstGroupsUsers.List(i)) Then
      lstGroupsUsers.Selected(i) = True
    Else
      lstGroupsUsers.Selected(i) = False
    End If
  Next
  mbSettingUser = False
  
  Exit Sub
  
GSErr:
  ShowError
  mbSettingUser = False
End Sub

Private Sub lstGroupsUsers_Click()
  On Error GoTo GUErr
  
  Dim usr As User
  
  If mbSettingUser Then Exit Sub
  
  If lstGroups.ListIndex < 0 Then
    Beep
    MsgBox MSG2
    Exit Sub
  End If
  
  If lstGroupsUsers.Selected(lstGroupsUsers.ListIndex) Then
    'agrega el usuario al grupo
    Set usr = gwsMainWS.CreateUser(lstGroupsUsers.Text)
    gwsMainWS.Groups(lstGroups.Text).Users.Append usr
    gwsMainWS.Users(lstGroupsUsers.Text).Groups.Refresh
  Else
    'elimina el usuario del grupo
    gwsMainWS.Groups(lstGroups.Text).Users.Delete lstGroupsUsers.Text
    gwsMainWS.Users(lstGroupsUsers.Text).Groups.Refresh
  End If
  
  Exit Sub
  
GUErr:
  ShowError
End Sub

Private Sub lstTablesQuerys_Click()
  SetPermissions False
End Sub

Private Sub lstUsers_Click()
  On Error GoTo USErr

  Dim i As Integer
  
  mbSettingUser = True
  For i = 0 To lstUsersGroups.ListCount - 1
    If IsMemberOf(lstUsersGroups.List(i), lstUsers.Text) Then
      lstUsersGroups.Selected(i) = True
    Else
      lstUsersGroups.Selected(i) = False
    End If
  Next
  mbSettingUser = False
  
  'muestra los permisos
  SetPermissions False
  
  Exit Sub
  
USErr:
  ShowError
  mbSettingUser = False
End Sub

Private Function IsMemberOf(rsGrp As String, rsUsr As String) As Integer
  Dim usr As User
  Dim grp As Group
  Dim i As Integer
  
  Set usr = gwsMainWS.Users(rsUsr)
  
  For Each grp In usr.Groups
    If grp.Name = rsGrp Then
      IsMemberOf = True
      Exit Function
    End If
  Next
  IsMemberOf = False
End Function

Private Sub lstUsersGroups_Click()
  On Error GoTo UGErr
  
  Dim grp As Group
  
  If mbSettingUser Then Exit Sub
  
  If lstUsers.ListIndex < 0 Then
    Beep
    MsgBox MSG4
    Exit Sub
  End If
  
  If lstUsersGroups.Selected(lstUsersGroups.ListIndex) Then
    'agrega el grupo al usuario
    Set grp = gwsMainWS.CreateGroup(lstUsersGroups.Text)
    gwsMainWS.Users(lstUsers.Text).Groups.Append grp
    gwsMainWS.Groups(lstUsersGroups.Text).Users.Refresh
  Else
    'elimina el grupo del usuario
    gwsMainWS.Users(lstUsers.Text).Groups.Delete lstUsersGroups.Text
    gwsMainWS.Groups(lstUsersGroups.Text).Users.Refresh
  End If
  
  Exit Sub
  
UGErr:
  ShowError
End Sub

Private Sub optGroups_Click()
  picUsers.Visible = False
  picGroups.Visible = True
End Sub

Private Sub optUsers_Click()
  picGroups.Visible = False
  picUsers.Visible = True
End Sub

Private Sub SetPermissions(rbAssign As Integer)
  On Error GoTo SPErr
  
  Dim lPermissions As Long
  Dim lPermissions2 As Long
  Dim bUncommon As Integer    'indicador de selección múltiple para permisos comunes
  Dim nCnt As Integer         'cuenta de objetos seleccionados
  Dim sUserGroup As String
  Dim sObject As String
  Dim i As Integer
  
  mbSettingPerm = True
  
  If rbAssign Then
    'determina los permisos establecidos y los combina con un Or
    If chkReadDesign.Value = vbUnchecked Then
      lPermissions = 0
    Else
      If chkAdminister.Value = vbChecked Then
        'los establece todos
        lPermissions = dbSecFullAccess Or _
                       dbSecReadDef Or _
                       dbSecWriteDef Or _
                       dbSecRetrieveData Or _
                       dbSecReplaceData Or _
                       dbSecInsertData Or _
                       dbSecDeleteData
      Else
        'los establece de uno en uno
        lPermissions = dbSecReadDef
        If chkModifyDesign.Value = vbChecked Then
          lPermissions = lPermissions Or dbSecWriteDef
        End If
        If chkReadData.Value = vbChecked Then
          lPermissions = lPermissions Or dbSecRetrieveData
        End If
        If chkUpdateData.Value = vbChecked Then
          lPermissions = lPermissions Or dbSecReplaceData
        End If
        If chkInsertData.Value = vbChecked Then
          lPermissions = lPermissions Or dbSecInsertData
        End If
        If chkDeleteData.Value = vbChecked Then
          lPermissions = lPermissions Or dbSecDeleteData
        End If
      End If
    End If
  End If
  
  'determina si es un usuario o un grupo
  If optUsers.Value Then
    'usuarios
    sUserGroup = lstUsers.Text
  Else
    'grupos
    sUserGroup = lstGroups.Text
  End If
  
  'establece u obtiene los permisos
  If lstTablesQuerys.ListIndex = -1 Then
    If mbLoading = False Then   'no emite errores al cargar el formulario
      Beep
      MsgBox MSG9
    End If
    Exit Sub
  End If
  
  'recorre la lista de objetos y procesa los objetos seleccionados
  For i = 0 To lstTablesQuerys.ListCount - 1
    If lstTablesQuerys.Selected(i) Then
      nCnt = nCnt + 1
      If lstTablesQuerys.ListIndex = 0 Then
        'tiene que ser <New Object>
        gdbCurrentDB.Containers("Tables").UserName = sUserGroup
        If rbAssign = False Then
          lPermissions = gdbCurrentDB.Containers("Tables").Permissions
        Else
          gdbCurrentDB.Containers("Tables").Permissions = lPermissions
        End If
        ShowOwner gdbCurrentDB.Containers("Tables")
        Set mobjCurrObject = gdbCurrentDB.Containers("Tables")
      Else
        sObject = StripConnect(lstTablesQuerys.List(i))
        'se ha seleccionado una tabla o una consulta
        gdbCurrentDB.Containers("Tables").Documents(sObject).UserName = sUserGroup
        If rbAssign = False Then
          lPermissions = gdbCurrentDB.Containers("Tables").Documents(sObject).Permissions
        Else
          gdbCurrentDB.Containers("Tables").Documents(sObject).Permissions = lPermissions
        End If
        ShowOwner gdbCurrentDB.Containers("Tables").Documents(sObject)
        Set mobjCurrObject = gdbCurrentDB.Containers("Tables").Documents(sObject)
      End If
      If nCnt > 1 Then
        'si hay más de 1, tienen que coincidir o establecemos el indicador
        If lPermissions <> lPermissions2 Then
          bUncommon = True
        End If
      End If
      'lo guarda para la próxima vez
      lPermissions2 = lPermissions
    End If
  Next
  
  If rbAssign = False Then
    If bUncommon Then
      'hay alguna diferencia, así que hay que ponerlos en gris
      chkReadDesign.Value = 2
      chkModifyDesign.Value = 2
      chkAdminister.Value = 2
      chkReadData.Value = 2
      chkUpdateData.Value = 2
      chkInsertData.Value = 2
      chkDeleteData.Value = 2
    Else
      'había solo una o eran todos igual
      'así que tenemos que establecerlos adecuadamente
      If (lPermissions And dbSecReadDef) = dbSecReadDef Then
        chkReadDesign.Value = vbChecked
      Else
        chkReadDesign.Value = vbUnchecked
      End If
      If (lPermissions And dbSecWriteDef) = dbSecWriteDef Then
        chkModifyDesign.Value = vbChecked
      Else
        chkModifyDesign.Value = vbUnchecked
      End If
      If (lPermissions And dbSecFullAccess) = dbSecFullAccess Then
        chkAdminister.Value = vbChecked
      Else
        chkAdminister.Value = vbUnchecked
      End If
      If (lPermissions And dbSecRetrieveData) = dbSecRetrieveData Then
        chkReadData.Value = vbChecked
      Else
        chkReadData.Value = vbUnchecked
      End If
      If (lPermissions And dbSecReplaceData) = dbSecReplaceData Then
        chkUpdateData.Value = vbChecked
      Else
        chkUpdateData.Value = vbUnchecked
      End If
      If (lPermissions And dbSecInsertData) = dbSecInsertData Then
        chkInsertData.Value = vbChecked
      Else
        chkInsertData.Value = vbUnchecked
      End If
      If (lPermissions And dbSecDeleteData) = dbSecDeleteData Then
        chkDeleteData.Value = vbChecked
      Else
        chkDeleteData.Value = vbUnchecked
      End If
    End If
  End If
  
  mbSettingPerm = False
  Exit Sub
  
SPErr:
  mbSettingPerm = False
  ShowError
End Sub

Private Sub ShowOwner(vObj As Object)
  On Error GoTo SOErr
  
  Dim i As Integer
  
  For i = 0 To cboOwners.ListCount - 1
    If cboOwners.List(i) = vObj.Owner Then
      mbSettingOwner = True
      cboOwners.ListIndex = i
      mbSettingOwner = False
      Exit For
    End If
  Next

  Exit Sub
  
SOErr:
  mbSettingOwner = True
  cboOwners.ListIndex = -1
  mbSettingOwner = False
  ShowError
End Sub
