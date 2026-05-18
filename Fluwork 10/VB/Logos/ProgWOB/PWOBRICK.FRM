VERSION 5.00
Begin VB.Form frmBricks 
   Caption         =   "Colección empleados - Casa de ladrillos"
   ClientHeight    =   3525
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   4995
   Begin VB.CommandButton cmdTrouble 
      Caption         =   "&Problema"
      Enabled         =   0   'False
      Height          =   465
      Left            =   3150
      TabIndex        =   8
      Top             =   2250
      Width           =   1545
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   285
      Left            =   3150
      TabIndex        =   9
      Top             =   2880
      Width           =   1545
   End
   Begin VB.CommandButton cmdListEmployees 
      Caption         =   "Act&ualizar lista"
      Height          =   285
      Left            =   3150
      TabIndex        =   7
      Top             =   1800
      Width           =   1545
   End
   Begin VB.CommandButton cmdDeleteEmployee 
      Caption         =   "&Eliminar"
      Height          =   285
      Left            =   3150
      TabIndex        =   6
      Top             =   1440
      Width           =   1545
   End
   Begin VB.CommandButton cmdAddEmployee 
      Caption         =   "&Agregar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   285
      Left            =   3150
      TabIndex        =   5
      Top             =   1080
      Width           =   1545
   End
   Begin VB.ListBox lstEmployees 
      Height          =   1845
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   2715
   End
   Begin VB.TextBox txtSalary 
      Height          =   285
      Left            =   2700
      TabIndex        =   3
      Top             =   450
      Width           =   1995
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "&Sueldo"
      Height          =   195
      Left            =   2700
      TabIndex        =   2
      Top             =   180
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "&Nombre"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2265
   End
End
Attribute VB_Name = "frmBricks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sbMain As New SmallBusiness3

Private Sub cmdAddEmployee_Click()
    Dim empNew As Employee
    If Not IsNumeric(txtSalary) Then
        MsgBox "El sueldo no es válido."
        With txtSalary
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    ' El nuevo empleado sólo puede ser creado
    '   llamando al método Add de la
    '   colección Employees.
    Set empNew = sbMain.Employees.Add(txtName.Text, CDbl(txtSalary.Text))
    ' Agrega el nuevo empleado al cuadro de lista.
    With empNew
        lstEmployees.AddItem .ID & ", " & .Name & ", " & .Salary
    End With
    With lstEmployees
        ' Selecciona el elemento recién agregado.
        .ListIndex = .NewIndex
    End With
    txtName.Text = ""
    txtSalary.Text = ""
    txtName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteEmployee_Click()
    Dim lngDeletedItem As Long
    With lstEmployees
        lngDeletedItem = .ListIndex
        ' Comprueba que hay un empleado seleccionado.
        If .ListIndex > -1 Then
            ' El ID del empleado son los seis primeros caracteres de la línea.
            sbMain.Employees.Delete Left(.Text, 6)
            ' Elimina el elemento seleccionado.
            .RemoveItem .ListIndex
            If .ListCount = 0 Then
                ' Si la lista está vacía,
                '   no permite establecer una nueva
                '   selección.
                Exit Sub
            End If
            ' ¿El elemento eliminado estaba al final del
            '   cuadro de lista?  Si es así, su índice será
            '   mayor o igual que la cuenta de la lista...
            If .ListCount <= lngDeletedItem Then
                '   ...de modo que establece la selección actual
                '   al nuevo elemento último...
                .ListIndex = lngDeletedItem - 1
            Else
                '   ...si no, mantiene la selección en la
                '   misma posición física
                '   de la lista.
                .ListIndex = lngDeletedItem
            End If
        Else
            MsgBox "No hay empleados seleccionados."
        End If
    End With
End Sub

Private Sub cmdListEmployees_Click()
    Dim emp As Employee
    With lstEmployees
        .Clear
        For Each emp In sbMain.Employees
            With emp
                lstEmployees.AddItem .ID & ", " & .Name & ", " & .Salary
            End With
        Next
        If .ListCount <> 0 Then
            ' Si hay elementos en la lista,
            '   selecciona el primero.
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Establece la variable global oculta de
    '   este formulario a Nothing, para liberar
    '   sus recursos.
    Set frmBricks = Nothing
End Sub

Private Sub txtName_Change()
    Call EnableAddButton
End Sub

Private Sub txtSalary_Change()
    Call EnableAddButton
End Sub

Private Sub txtSalary_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48 To 57   ' Permite los dígitos
        Case 8      ' Permite el carácter de retroceso
        Case 46     ' Permite el punto
        Case Else
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub EnableAddButton()
    If (Len(txtName) > 0) And (Len(txtSalary) > 0) Then
        cmdAddEmployee.Enabled = True
    Else
        cmdAddEmployee.Enabled = False
    End If
End Sub

