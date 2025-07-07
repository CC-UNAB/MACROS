Attribute VB_Name = "Módulo2"
Option Explicit

'EXCELeINFO
'MVP Sergio Alejandro Campos
'http://www.exceleinfo.com
'https://www.youtube.com/user/sergioacamposh
'http://blogs.itpro.es/exceleinfo

Sub ExtraerCorreosDeOutlook()

Dim OutlookApp As Outlook.Application
Dim ONameSpace As Object
Dim MyFolder As Object
Dim OItem As Outlook.MailItem
Dim Fila As Integer
Dim FolderDescarga As String
Dim Adjuntos As Integer
Dim NombreArchivo1, NombreArchivo
Dim i As Integer
Dim Fecha1 As Date
Dim Fecha2 As Date

Set OutlookApp = New Outlook.Application
Set ONameSpace = OutlookApp.GetNamespace("MAPI")
Set MyFolder = ONameSpace.GetDefaultFolder(olFolderInbox)
'Set MyFolder = ONameSpace.Folders("susana.lefimil@unab.cl").Folders("Prueba")

FolderDescarga = ThisWorkbook.Path & "Extract"

Range(Range("A2"), ActiveCell.SpecialCells(xlLastCell)).ClearContents

Fila = 2
Fecha1 = "01/01/2020"
Fecha2 = "31/12/2020"

For Each OItem In MyFolder.Items

    If Int(OItem.ReceivedTime) >= Fecha1 And Int(OItem.ReceivedTime) <= Fecha2 Then

        Adjuntos = 0
        NombreArchivo1 = ""
        
        If OItem.Attachments.Count > 0 Then
        
            For i = 1 To OItem.Attachments.Count
                NombreArchivo = OItem.Attachments.Item(i).Filename
                OItem.Attachments.Item(i).SaveAsFile FolderDescarga & "" & NombreArchivo
                NombreArchivo1 = NombreArchivo & ", " & NombreArchivo1
                Adjuntos = Adjuntos + 1
            Next i
        End If
            Sheets("Hoja1").Cells(Fila, 1).Value = OItem.SenderEmailAddress
            Sheets("Hoja1").Cells(Fila, 2).Value = OItem.To
            Sheets("Hoja1").Cells(Fila, 3).Value = OItem.Subject
            Sheets("Hoja1").Cells(Fila, 4).Value = OItem.ReceivedTime
            Sheets("Hoja1").Cells(Fila, 5).Value = Adjuntos
            Sheets("Hoja1").Cells(Fila, 6).Value = NombreArchivo1
            Sheets("Hoja1").Cells(Fila, 7).Value = OItem.Body
    
        Fila = Fila + 1
        
    End If

Next OItem

Set OutlookApp = Nothing
Set ONameSpace = Nothing
Set MyFolder = Nothing

End Sub
