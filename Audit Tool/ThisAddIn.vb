Imports System.Xml
Imports System.Environment
Imports System.Windows.Forms
Imports System.IO.Compression
Imports System.IO

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim XMLDir As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer"
        Dim XML As String = XMLDir & "\BranchCodes.XML"
        If Not My.Computer.FileSystem.DirectoryExists(XMLDir) Then
            My.Computer.FileSystem.CreateDirectory(XMLDir)
        End If
        If Not My.Computer.FileSystem.FileExists(XML) Then
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True

            Using writer As XmlWriter = XmlWriter.Create(XML, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement("BranchCodes")
                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using
        End If
        If Password().Length = 0 Then
            Dim Password As Form = New Password
            Password.ShowDialog()
        End If
        'Barcode()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Function Password() As String
        Password = String.Empty
        Dim XMLDir As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer"
        Dim XML As String = XMLDir & "\Password.XML"
        If Not My.Computer.FileSystem.DirectoryExists(XMLDir) Then
            My.Computer.FileSystem.CreateDirectory(XMLDir)
        End If
        If Not My.Computer.FileSystem.FileExists(XML) Then
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True

            Using writer As XmlWriter = XmlWriter.Create(XML, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement("Settings")
                writer.WriteStartElement("Password")
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using
        End If
        Dim XMLDoc As New XmlDocument
        XMLDoc.Load(XML)
        For Each Node As XmlNode In XMLDoc.DocumentElement.ChildNodes
            Password = Node.InnerText
        Next
    End Function

End Class
