Private WithEvents oItems As Outlook.Items

' Configuración de inicio de la aplicación y opciones de almacenamiento
Private Sub Application_Startup()
    On Error GoTo ErrorHandler

    ' Inicialización del sistema de archivos y carpetas
    Dim sistema As Object
    Dim rutaCarpetaRaiz As String
    Dim nombreCarpetaAdjuntos As String
    rutaCarpetaRaiz = "C:\Users\goreg\Documents"
    nombreCarpetaAdjuntos = "DatosAdjuntos"

    ' Crear el objeto FileSystemObject
    Set sistema = CreateObject("Scripting.FileSystemObject")

    ' Crear la carpeta principal de adjuntos
    If Not sistema.FolderExists(rutaCarpetaRaiz & "\" & nombreCarpetaAdjuntos) Then
        sistema.CreateFolder rutaCarpetaRaiz & "\" & nombreCarpetaAdjuntos
    End If

    ' Configuración de opción de guardado en borrador
    Dim cuenta As Outlook.Account
    Dim borradorOpciones As Outlook.MailItem
    Dim borradorFecha As Outlook.MailItem
    Dim borradorEjecucion As Outlook.MailItem
    Dim ejecution As String
    Dim valCapetaRaiz As Boolean
    
    ejecution = "0"
    valCarpetaRaiz = False
    Dim sUserChoice As String
    Dim fechaUltimaEjecucion As Date

    For Each cuenta In Outlook.Application.Session.Accounts
        ' Borrador de configuración de opciones
        Set borradorOpciones = cuenta.Session.GetDefaultFolder(olFolderDrafts).Items.Find("[Subject] = 'IMPORTANTE CONFIG DB'")
        
        ' Si el borrador de opciones no existe, crearlo y solicitar la opción de guardado
        If borradorOpciones Is Nothing Then
            Set borradorOpciones = Application.CreateItem(olMailItem)
            borradorOpciones.Subject = "IMPORTANTE CONFIG DB"
            sUserChoice = InputBox("Escaner PDF JSON" & vbCrLf & " " & vbCrLf & "Ingrese la opción de guardado:" & vbCrLf & "1 - Guardar por remitente" & vbCrLf & "2 - Guardar por tipo de archivo")
            If sUserChoice = "" Then
                borradorOpciones.Body = 0
                borradorOpciones.Save
            Else
                borradorOpciones.Body = sUserChoice
                borradorOpciones.Save
            End If
            
        Else
            sUserChoice = borradorOpciones.Body
        End If
        
        ' Verifica si el borrador de ejecución existe y créalo si no
        Set borradorEjecucion = cuenta.Session.GetDefaultFolder(olFolderDrafts).Items.Find("[Subject] = 'IMPORTANTE EJECUCION DB'")
        
        If borradorEjecucion Is Nothing Then
            Set borradorEjecucion = Application.CreateItem(olMailItem)
            borradorEjecucion.Subject = "IMPORTANTE EJECUCION DB"
            borradorEjecucion.Body = ejecution
            borradorEjecucion.Save
        End If

        ' Borrador para la última fecha de ejecución
        Set borradorFecha = cuenta.Session.GetDefaultFolder(olFolderDrafts).Items.Find("[Subject] = 'IMPORTANTE FECHA DB'")
        
        ' Si el borrador de fecha no existe, crearlo y asignar la fecha actual
        If borradorFecha Is Nothing Then
            Set borradorFecha = Application.CreateItem(olMailItem)
            borradorFecha.Subject = "IMPORTANTE FECHA DB"
            borradorFecha.Body = Now
            borradorFecha.Save
            fechaUltimaEjecucion = Now
        Else
            fechaUltimaEjecucion = CDate(borradorFecha.Body)
        End If
        
        ' Dependiendo de la opción seleccionada, ejecutar el procesamiento adecuado
        If sUserChoice = 1 Then
            Set oItems = Outlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
            ProcesarCorreosPorRemitente cuenta, sistema, nombreCarpetaAdjuntos, rutaCarpetaRaiz, fechaUltimaEjecucion, borradorFecha, borradorEjecucion
        ElseIf sUserChoice = 2 Then
            ProcesarCorreosPorTipo cuenta, sistema, nombreCarpetaAdjuntos, rutaCarpetaRaiz, fechaUltimaEjecucion, borradorFecha, borradorEjecucion
        Else
            MsgBox "ESCANER PDF JSON " & vbCrLf & " " & vbCrLf & "Opción no válida o carpeta de destino no encontrada. Por favor, reinicie Outlook y seleccione una opción. De lo contrario, no se podrán realizar búsquedas para su almacenamiento posterior."
            borradorOpciones.Delete
            borradorFecha.Delete
            borradorEjecucion.Delete
        End If
    Next

    Exit Sub

ErrorHandler:
    MsgBox "ESCANER PDF JSON " & vbCrLf & " " & vbCrLf & "Se le informa que no se llevarán a cabo acciones de escaneo y guardado. Una vez que esté listo, por favor reinicie la aplicación y seleccione una opción."
    borradorOpciones.Delete
    borradorFecha.Delete
    borradorEjecucion.Delete
End Sub

' Sub para procesar correos por remitente
Private Sub ProcesarCorreosPorRemitente(cuenta As Outlook.Account, sistema As Object, nombreCarpetaAdjuntos As String, rutaCarpetaRaiz As String, fechaUltimaEjecucion As Date, borradorFecha As Outlook.MailItem, borradorEjecucion As Outlook.MailItem)
    Dim correo As Object
    Dim archivoAdjunto As Attachment
    Dim listaCorreos As Collection
    Set listaCorreos = New Collection

    ' Valida su primera ejecución
    If borradorEjecucion.Body = 0 Then
    
        For Each correo In cuenta.Session.GetDefaultFolder(olFolderInbox).Items
            If TypeName(correo) = "MailItem" Then
                listaCorreos.Add correo
            End If
        Next correo
        borradorEjecucion.Body = 1
        borradorEjecucion.Save
    Else
        For Each correo In cuenta.Session.GetDefaultFolder(olFolderInbox).Items
            If TypeName(correo) = "MailItem" And correo.ReceivedTime > fechaUltimaEjecucion Then
                listaCorreos.Add correo
            End If
        Next correo
    End If

    ' Procesar cada correo en la colección temporal
    For Each correo In listaCorreos
        Dim tienePDF As Boolean, tieneJSON As Boolean
        tienePDF = False
        tieneJSON = False

        For Each archivoAdjunto In correo.attachments
            If archivoAdjunto.FileName Like "*.pdf" Then
                tienePDF = True
            ElseIf archivoAdjunto.FileName Like "*.json" Then
                tieneJSON = True
            End If
            If tienePDF And tieneJSON Then Exit For
        Next archivoAdjunto

        ' Crear carpetas solo si hay archivos PDF o JSON
        If tienePDF Or tieneJSON Then
            GuardarAdjuntosPorRemitente sistema, correo, nombreCarpetaAdjuntos, rutaCarpetaRaiz
        End If
    Next correo

    ' Actualizar la última fecha de ejecución en el borrador de fecha
    borradorFecha.Body = CStr(Now)
    borradorFecha.Save
End Sub

' Sub para guardar archivos adjuntos por remitente
Private Sub GuardarAdjuntosPorRemitente(sistema As Object, correo As Outlook.MailItem, nombreCarpetaAdjuntos As String, rutaCarpetaRaiz As String)
    Dim rutaCarpetaDominio As String
    rutaCarpetaDominio = rutaCarpetaRaiz & "\" & nombreCarpetaAdjuntos & "\" & correo.SenderEmailAddress

    ' Crear carpeta del remitente si no existe
    If Not sistema.FolderExists(rutaCarpetaDominio) Then
        sistema.CreateFolder rutaCarpetaDominio
    End If

    ' Crear subcarpetas PDF y JSON
    If Not sistema.FolderExists(rutaCarpetaDominio & "\pdf") Then
        sistema.CreateFolder rutaCarpetaDominio & "\pdf"
    End If
    If Not sistema.FolderExists(rutaCarpetaDominio & "\json") Then
        sistema.CreateFolder rutaCarpetaDominio & "\json"
    End If

    ' Guardar archivos adjuntos en las carpetas correspondientes
    For Each archivoAdjunto In correo.attachments
        If archivoAdjunto.FileName Like "*.pdf" Then
            archivoAdjunto.SaveAsFile rutaCarpetaDominio & "\pdf\" & archivoAdjunto.FileName
        ElseIf archivoAdjunto.FileName Like "*.json" Then
            archivoAdjunto.SaveAsFile rutaCarpetaDominio & "\json\" & archivoAdjunto.FileName
        End If
    Next archivoAdjunto
End Sub

' Sub para procesar correos por tipo de archivo
Private Sub ProcesarCorreosPorTipo(cuenta As Outlook.Account, sistema As Object, nombreCarpetaAdjuntos As String, rutaCarpetaRaiz As String, fechaUltimaEjecucion As Date, borradorFecha As Outlook.MailItem, borradorEjecucion As Outlook.MailItem)
    Dim correo As Object
    Dim archivoAdjunto As Attachment
    Dim dataCorreosTemp As Collection

    Set dataCorreosTemp = New Collection

    ' Validar primera ejecución
    If borradorEjecucion.Body = 0 Then
    
        For Each correo In cuenta.Session.GetDefaultFolder(olFolderInbox).Items
            If TypeName(correo) = "MailItem" Then
                dataCorreosTemp.Add correo
            End If
        Next correo
        borradorEjecucion.Body = 1
        borradorEjecucion.Save
    Else
        For Each correo In cuenta.Session.GetDefaultFolder(olFolderInbox).Items
            If TypeName(correo) = "MailItem" And correo.ReceivedTime > fechaUltimaEjecucion Then
                dataCorreosTemp.Add correo
            End If
        Next correo
    End If

    ' Procesar cada correo en la colección temporal
    For Each correo In dataCorreosTemp
        Dim tienePDF As Boolean, tieneJSON As Boolean
        tienePDF = False
        tieneJSON = False

        For Each archivoAdjunto In correo.attachments
            If archivoAdjunto.FileName Like "*.pdf" Then
                tienePDF = True
            ElseIf archivoAdjunto.FileName Like "*.json" Then
                tieneJSON = True
            End If
            If tienePDF And tieneJSON Then Exit For
        Next archivoAdjunto

        ' Guardar adjuntos solo si hay archivos PDF o JSON
        If tienePDF Or tieneJSON Then
            GuardarAdjuntosPorTipo sistema, correo, nombreCarpetaAdjuntos, rutaCarpetaRaiz
        End If
    Next correo

    ' Actualizar la última fecha de ejecución en el borrador de fecha
    borradorFecha.Body = CStr(Now)
    borradorFecha.Save
End Sub

' Sub para guardar archivos adjuntos por tipo de archivo
Private Sub GuardarAdjuntosPorTipo(sistema As Object, correo As Outlook.MailItem, nombreCarpetaAdjuntos As String, rutaCarpetaRaiz As String)
    Dim rutaCarpetaPDF As String
    Dim rutaCarpetaJSON As String
    rutaCarpetaPDF = rutaCarpetaRaiz & "\" & nombreCarpetaAdjuntos & "\pdf"
    rutaCarpetaJSON = rutaCarpetaRaiz & "\" & nombreCarpetaAdjuntos & "\json"

    ' Crear las carpetas PDF y JSON si no existen
    If Not sistema.FolderExists(rutaCarpetaPDF) Then
        sistema.CreateFolder rutaCarpetaPDF
    End If
    If Not sistema.FolderExists(rutaCarpetaJSON) Then
        sistema.CreateFolder rutaCarpetaJSON
    End If

    ' Guardar archivos adjuntos en las carpetas correspondientes
    For Each archivoAdjunto In correo.attachments
        If archivoAdjunto.FileName Like "*.pdf" Then
            archivoAdjunto.SaveAsFile rutaCarpetaPDF & "\" & archivoAdjunto.FileName
        ElseIf archivoAdjunto.FileName Like "*.json" Then
            archivoAdjunto.SaveAsFile rutaCarpetaJSON & "\" & archivoAdjunto.FileName
        End If
    Next archivoAdjunto
End Sub
