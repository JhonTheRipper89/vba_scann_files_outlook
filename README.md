# VBA Outlook

##
Este  código  VBA  para  Outlook  está  diseñado  para  escanear  correos  electrónicos  y 
almacenar automáticamente los archivos adjuntos PDF y JSON en carpetas específicas 
en el sistema de archivos. Su configuración permite que el usuario elija cómo organizar 
los archivos adjuntos, ya sea por remitente o por tipo de archivo, y conserva un historial 
de la última fecha de ejecución para evitar procesamientos repetidos
##

## Requisitos previos

1. Microsoft Outlook (classic) instalado en tu equipo.

## Paso 1: Configuración de macros 
1. Abre Microsoft Outlook. 
2. Dirígete a Archivo > Opciones. 
3. Selecciona Centro de confianza > Configuración del Centro de confianza. 
4. Dentro de la configuración, selecciona Configuración de macros. 
5. Marca la opción Habilitar todas las macros. 

## Paso 2: Activar el Modo Desarrollador 
1. Ve a Archivo > Opciones. 
2. Selecciona Personalizar cinta de opciones.
3. Marca la opción Programador o Desarrollador para habilitar esta pestaña. 

## Paso 3: Activar Microsoft Scripting Runtime 
1. En la pestaña Programador, selecciona Visual Basic
2. En la ventana que se abre, ve a Herramientas > Referencias.
3. Marca la opción Microsoft Scripting Runtime. 

## Paso 4: Agregar el código 
1. En Programador > Visual Basic, haz doble clic en ThisOutlookSession en el panel de 
la izquierda. 
2. Pega el código proporcionado en el archivo de la macro en esta ventana. 

## Paso 5: Configuración de la ruta de guardado
1. Dentro del código, localiza la sección donde se especifica la ruta de almacenamiento. 
2. Cambia la ruta según el directorio en el que deseas almacenar los archivos adjuntos. 
3. Guarda los cambios con Ctrl+s 

¡Todo estará listo para comenzar! 






[Wv]( http://www.linkedin.com/in/william-ventura-aa66bb324 )