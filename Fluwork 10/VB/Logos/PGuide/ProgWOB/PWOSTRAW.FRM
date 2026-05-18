VERSION 5.00
Begin VB.Form frmStraw 
   Caption         =   "Colección empleados - Casa de paja"
   ClientHeight    =   3525
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3525
   ScaleWidth      =   4995
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdTrouble 
      Caption         =   "&Problema"
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
      Sorted          =   -1  'True
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
Attribute VB_Name = "frmStraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sbMain As New SmallBusiness1

Private Sub cmdAddEmployee_Click()
    Dim empNew As New Employee
    If Not IsNumeric(txtSalary) Then
        MsgBox "Sueldo no tiene un valor válido."
        ' Establece el enfoque en el campo salario,
        '   y selecciona todo el texto.
        With txtSalary
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        Exit Sub
    End If
    With empNew
        .ID = sbMain.NewEmployeeID
        .Name = txtName.Text
        .Salary = CDbl(txtSalary.Text)
        sbMain.Employees.Add empNew, .ID
        lstEmployees.AddItem .ID & ", " & .Name & ", " & .Salary
        With lstEmployees
            ' Selecciona el elemento recién agregado.
            .ListIndex = .NewIndex
        End With
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
        ' Asegura que hay un empleado seleccionado.
        If .ListIndex > -1 Then
            ' El ID del empleado son los seis primeros caracteres de la línea.
            sbMain.Employees.Remove Left(lstEmployees.Text, 6)
            ' Elimina el elemento seleccionado.
            .RemoveItem .ListIndex
            If .ListCount = 0 Then
                ' Si la lista está vacía,
                '   no intenta establecer una nueva
                '   selección.
                Exit Sub
            End If
            ' ¿El elemento seleccionado estaba al final del
            '   cuadro de lista?  Si lo está, su índice será mayor
            '   o igual que la cuenta de la lista ...
            If .ListCount <= lngDeletedItem Then
                '   ... de modo que establece la selección al
                '   nuevo elemento inferior ...
                .ListIndex = lngDeletedItem - 1
            Else
                '   ...si no, mantiene la selección
                '   en la misma posición física de
                '   la lista.
                .ListIndex = lngDeletedItem
            End If
        Else
            MsgBox "No hay ningún empleado seleccionado."
        End If
    End With
End Sub

Private Sub cmdListEmployees_Click()
    Dim emp As Employee
    With lstEmployees
        .Clear
        For Each emp In sbMain.Employees
            .AddItem emp.ID & ", " & emp.Name & ", " & emp.Salary
            ' Después de presionar el botón Problema, el clic en
            '   Actualizar genera un error de no coincidencia de tipos
            '   (en la instrucción For Each, si el elemento no válido
            '   es el primero de la lista, o en la
            '   instrucción Next) cuando Visual Basic intenta poner
            '   la referencia a frmStraw en la variable
            '   de iteración emp. Para continuar la ejecución,
            '   arrastre la flecha de ejecución amarilla a
            '   End Sub (o haga clic en End Sub y presione
            '   Ctrl+F9), después presione F5.
        Next
        ' Cuando se detenga aquí, lea el comentario anterior.
        '
        If .ListCount <> 0 Then
            ' Si hay elementos en la lista,
            '   selecciona el primero.
            .ListIndex = 0
        End If
    End With
End Sub

Private Sub cmdTrouble_Click()
    ' Como el objeto Collection acepta cualquier
    '   objeto, un error del código puede colocar
    '   un objeto no válido en la colección.
    sbMain.Employees.Add Me
    MsgBox "Se ha agregado a la colección una referencia al formulario de entrada de datos. Presione Actualizar lista para ver el error que lo produce."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Establece todas las referencias a este formulario a
    '   Nothing, para liberar sus recursos.
    '   Esto significa dos cosas:
    '   (1) Establece la variable global oculta
    '       del formulario a Nothing:
    Set frmStraw = Nothing
    '   (2) Borra el objeto colección,
    '       porque el botón Problema
    '       pone una referencia del formulario
    '       en la colección, creando una
    '       referencia circular (sbMain
    '       tiene una referencia a Employees,
    '       que tiene una referencia al
    '       formulario, que tiene una referencia
    '       a sbMain) que mantiene vivos todos los
    '       objetos.
    Set sbMain.Employees = Nothing
    '
    ' Por supuesto, el que no podamos destruir
    '   la colección Employees del objeto
    '    SmallBusiness de esta manera, es un error;
    '   pero después de todo, la Casa de paja
    '   NO es la manera de hacer las cosas.
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
