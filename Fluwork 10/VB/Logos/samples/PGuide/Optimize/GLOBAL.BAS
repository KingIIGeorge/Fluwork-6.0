Attribute VB_Name = "Global"
Option Explicit
Public Const VBVer = "5.0"
Public Const LastAppUpdate = "4/9/96"

Public Const WM_SYSCOMMAND = &H112
Public Const SC_RESTORE = &HF120
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const WM_SETREDRAW = &HB

' Llamadas al API de Windows para crear la ventana Topmost
Declare Function SetWindowPos Lib "User32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Declare Function FindWindow Lib "User32" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function SendMessage Lib "User32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Type BookAuthor
    ID As Integer
    Name As String * 25
End Type

Dim arDemos(20) As String

Public Function SetRedraw(hWnd As Long, lState As Long) As Long
  
  Dim lRetVal As Long
  
  On Error GoTo SetRedraw_Error
  ' Establezca iState = 0 para desactivar la actualización automática
  ' Establezca iState = 1 para activar la actualización automática
  lRetVal = SendMessage(hWnd, WM_SETREDRAW, lState, 0)
  If lRetVal <> False Then GoTo SetRedraw_Error
  
  SetRedraw = True
  Exit Function
  
SetRedraw_Error:
  SetRedraw = lRetVal
  Exit Function
  
End Function

Public Function SetTopWindow(hWnd As Long, bState As Boolean) As Boolean
      
  If bState = True Then 'Pone la ventana en primer plano
    SetTopWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  ElseIf bState = False Then ' Desactiva el indicador TopMost
    SetTopWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    Debug.Print "bState desconocido."
    SetTopWindow = False
  End If
  
End Function

' Centra el objeto Form que recibe.
Sub CenterMe(frmForm As Form)
  frmForm.Left = (Screen.Width - frmForm.Width) / 2
  frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub

Sub Main()
  Dim iCount As Integer
  Dim iLimit As Integer
  Dim tvwSample As TreeView
  Dim nodRoot As Node
  Dim nodX As Node
  Dim sLastParent As String
  Dim ilImages As ListImages
  Dim iNextParentImage As Integer
  
  Set tvwSample = frmExplore.tvExample
  Set ilImages = frmExplore.ilExplore.ListImages
  
  frmExplore.Show
  
  ' Presenta las categorías de velocidad
  arDemos(0) = "Algoritmos"
  arDemos(1) = "Velocidad de visualización"
  arDemos(2) = "Dibujar imagen"
  arDemos(3) = "Velocidad de visualización"
  arDemos(4) = "Gráficos"
  arDemos(5) = "Velocidad de visualización"
  
  ' Categorías de velocidad real
  arDemos(6) = "Cadenas y archivos"
  arDemos(7) = "Velocidad real"
  arDemos(8) = "Optimización de código"
  arDemos(9) = "Velocidad real"
  arDemos(10) = "Tipos numéricos"
  arDemos(11) = "Velocidad real"
  
  ' Categorías de uso de recursos
  arDemos(12) = "Picture frente a Image"
  arDemos(13) = "Uso de recursos"
  arDemos(14) = "Destruir objetos de formularios"
  arDemos(15) = "Uso de recursos"
  
  ' Categorías de velocidad aparente
  arDemos(16) = "Pantalla inicial"
  arDemos(17) = "Velocidad aparente"
    
  ' Categorías de colección
  arDemos(18) = "Colecciones"
  arDemos(19) = "Colecciones"
  
  Set nodRoot = tvwSample.Nodes.Add(, , "Root", "Ejemplos de optimización")
  nodRoot.Image = 1
  iNextParentImage = 2
  'Configura el primer tipo de Demo
  'Crea un nuevo tipo Demo a partir de Root y hace que este nodo sea hijo suyo.
  iCount = 0
  sLastParent = arDemos(iCount + 1)
  Set nodX = tvwSample.Nodes.Add("Root", tvwChild, sLastParent, arDemos(iCount + 1))
  nodX.Image = iNextParentImage  'Presenta la velocidad
  iNextParentImage = iNextParentImage + 1
  'Hace que este nodo sea hijo del último tipo Demo
  Set nodX = tvwSample.Nodes.Add(sLastParent, tvwChild, , arDemos(iCount))
  nodX.Image = 7
  iLimit = UBound(arDemos)
  For iCount = 2 To iLimit - 1 Step 2
    If arDemos(iCount + 1) <> arDemos(iCount - 1) Then
      nodX.EnsureVisible  ' Asegura que todos los hijos sean visibles en el último grupo
      'Crea un nuevo tipo Demo a partir de Root y hace que este nodo sea hijo suyo.
      sLastParent = arDemos(iCount + 1)
      Set nodX = tvwSample.Nodes.Add("Root", tvwChild, sLastParent, arDemos(iCount + 1))
      nodX.Image = iNextParentImage
      iNextParentImage = iNextParentImage + 1
    End If
    'Hace que este nodo sea hijo del último tipo Demo
    Set nodX = tvwSample.Nodes.Add(sLastParent, tvwChild, , arDemos(iCount))
    nodX.Image = 7
  Next iCount
    
  nodX.EnsureVisible
  
  ' Destruye todos los objetos que hemos creado.
  Set nodX = Nothing
  Set nodRoot = Nothing
  Set ilImages = Nothing
  Set tvwSample = Nothing

End Sub

Sub PosForm(frmForm As Form)
  frmForm.Left = frmExplore.Left + frmExplore.Width
  frmForm.Top = frmExplore.Top
End Sub


