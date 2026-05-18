Directorio ODBC
===============

Es posible usar el directorio \VB\ODBC para redistribuir ODBC y los controladores ODBC que se
suministran con Visual Basic 5.0 en su aplicación. El Asistente para instalar de Visual Basic
5.0 no es compatible con la instalación y distribución de ODBC o de los controladores ODBC que
se encuentran en su equipo. Este directorio se suministra para ayudarle a distribuir e instalar
adecuadamente todos los componentes necesarios para usar ODBC con su aplicación.

Contenido
---------
La instalación que se proporciona es una instalación estándar que usan otras aplicaciones de 
Microsoft, como Visual Basic 5.0. Instalar tiene un modo de mantenimiento después de la
instalación que se puede usar mediante el Panel de control (Agregar o quitar programas). La
instalación ODBC instala los siguientes componentes:

- Núcleo ODBC
- Controlador ODBC de Microsoft SQL Server
- Controlador ODBC de Microsoft Oracle
- Microsoft ODBC Desktop Driver Pack

Al usuario final se le presenta la opción de instalar todos los controladores ODBC o bien 
puede elegir los controladores que desea instalar mediante la opción Instalación personalizada.

Distribución
------------
Puede distribuir el contenido del directorio tal como está o puede crear discos de instalación
para el contenido del directorio.

Crear discos de instalación
---------------------------
Para distribuir ODBC en discos, necesitará tres discos con formato de 1,44 MB (o superior).
Puede copiar el contenido en los discos y ejecutar setup.exe desde el disco 1.  

Disco 1
   Setup.exe
   Setup.ini
   Setup.lst
   Setup.tdf
   Odbcstp.inf
   Odbc-1.cab

Disco 2
   Odbc2-cab

Disco 3
   Odbc3-cab

Ejecutar la instalación de ODBC desde la red
--------------------------------------------
Si está distribuyendo su aplicación en una red, puede copiar el contenido de este directorio a
una unidad compartida y ejecutar setup.exe desde esa ubicación.

Transferir componentes ODBC en Internet
---------------------------------------
La instalación de ODBC se ha comprimido en un archivo .cab que se puede usar para transferir
automáticamente e instalar ODBC mediante la transferencia de componentes. Es posible encontrar
el archivo Cabinet (ODBC30.Cab) en el directorio \Tools\Cabinets y en
http://activex.microsoft.com/controls/vb5.

Las siguientes entradas de ejemplo de archivo .INF muestran cómo incluir ODBC en su paquete de
transferencia.

[version]
signature="$CHICAGO$"
AdvancedINF=2.0

;Usa el controlador ODBC que desea instalar
[Add.Code]
sqlsrv32.dll=sqlsrv32.dll

;Sustituye FileVersion con la versión adecuada del controlador ODBC
[sqlsrv32.dll]
hook=ODBC30.cab_Installer
FileVersion=X,X,X,X

[ODBC30.cab_Installer]
file-win32-x86=ODBC30.cab
run=%EXTRACT_DIR%\ODBC30.exe
