Attribute VB_Name = "modMT"
Option Explicit

' Este módulo ilustra la forma en que Visual Basic
'   crea una instancia separada de los datos
'   globales para cada subroceso. Así, la variable
'   glngGlobalData tendrá un valor distinto para
'   cada subroceso que inicie Visual Basic.
'
' Este hecho se aprovecha para mantener la cuenta
'   de objetos Coffee en cada subroceso: En el evento
'   Initialize de cada objeto Coffee, se agrega uno
'   a glngGlobalData; y en su evento Terminate se
'   le resta uno.
'
' Cualquier objeto Coffee puede entonces averiguar
'   cuántos objetos Coffee hay en su subroceso,
'   observando glngGlobalData. Los clientes también
'   pueden saberlo llamando al método
'   NumberOnThread, que devuelve glngGlobalData.
'
' Cuando se ejecuta MTCoffee en el entorno de
'   desarrollo, en el que sólo hay un subroceso,
'   NumberOnThread es el número total de objetos
'   Coffee. Cuando MTCoffee se compila seleccionando
'   Subroceso por objeto, la cuenta será uno (1),
'   para cada objeto Coffee, salvo si llama a
'   GetCoffeeOnSameThread para crear un segundo
'   objeto Coffee en el mismo subroceso.
'
' Cuando se compila MTCoffee con pool de subrocesos,
'   NumberOnThread será mayor que uno en algún
'   subroceso siempre que el número de objetos
'   Coffee activos sea mayor que el número de
'   subrocesos del pool.
'
' Este tema se trata en "Escalabilidad y subrocesos
'   múltiples", en los Libros en pantalla.

Public glngGlobalData As Long

