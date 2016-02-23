Imports System.Security.Cryptography
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Environment
Imports System.Xml

Public Class Password
    Private Sub Password_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GUISetup()
    End Sub

    Sub GUISetup()
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Dim Padding As Integer = 15

        Dim LabelOne As New Label
        LabelOne.Width = 200
        LabelOne.Location = New Point(Padding, Padding)
        LabelOne.BackColor = Color.Transparent
        Me.Controls.Add(LabelOne)

        Dim TextOne As New TextBox
        TextOne.Name = "TextOne"
        TextOne.Location = New Point(Padding, LabelOne.Location.Y + LabelOne.Height)
        TextOne.Width = 200
        TextOne.PasswordChar = "*"
        TextOne.TabIndex = 0
        Me.Controls.Add(TextOne)

        Dim LabelTwo As New Label
        LabelTwo.Width = 200
        LabelTwo.Location = New Point(Padding, TextOne.Location.Y + TextOne.Height + Padding)
        LabelTwo.BackColor = Color.Transparent
        Me.Controls.Add(LabelTwo)

        Dim TextTwo As New TextBox
        TextTwo.Name = "TextTwo"
        TextTwo.Location = New Point(Padding, LabelTwo.Location.Y + LabelTwo.Height)
        TextTwo.Width = 200
        TextTwo.PasswordChar = "*"
        TextTwo.TabIndex = 1
        Me.Controls.Add(TextTwo)

        Dim Submit As New Button
        Submit.Text = "Submit"
        Submit.Location = New Point(Padding, TextTwo.Location.Y + TextTwo.Height + Padding)
        Submit.Width = 200
        Me.Controls.Add(Submit)

        If Password.Length = 0 Then
            Me.Text = "Password for Audit Analyzer"
            LabelOne.Text = "Password"
            LabelTwo.Text = "Confirm Password"
            AddHandler Submit.Click, AddressOf NewPassword
        Else
            Dim LabelThree As New Label
            LabelThree.Width = 200
            LabelThree.Location = New Point(Padding, TextTwo.Location.Y + TextTwo.Height + Padding)
            LabelThree.BackColor = Color.Transparent
            Me.Controls.Add(LabelThree)

            Dim TextThree As New TextBox
            TextThree.Name = "TextThree"
            TextThree.Location = New Point(Padding, LabelThree.Location.Y + LabelThree.Height)
            TextThree.Width = 200
            TextThree.PasswordChar = "*"
            TextThree.TabIndex = 2
            Me.Controls.Add(TextThree)

            Me.Text = "Change Password"
            LabelOne.Text = "Current Password"
            LabelTwo.Text = "New Password"
            LabelThree.Text = "Confirm Password"
            Submit.Location = New Point(Padding, TextThree.Location.Y + TextThree.Height + Padding)
            Submit.TabIndex = 3
            AddHandler Submit.Click, AddressOf ChangePassword
        End If

        Me.Size = New Size(Submit.Width + Padding * 2 + Me.Width - Me.ClientSize.Width, Submit.Location.Y + Submit.Height + Padding + Me.Height - Me.ClientSize.Height)
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
    End Sub

    Sub NewPassword()
        Dim TextOne As TextBox = Me.Controls("TextOne")
        Dim TextTwo As TextBox = Me.Controls("TextTwo")
        If String.IsNullOrEmpty(TextOne.Text) Then
            MsgBox("Password cannot be left blank.")
        ElseIf Not TextOne.Text = TextTwo.Text Then
            MsgBox("Passwords do not match.")
        ElseIf TextOne.Text.Length < 6 Then
            MsgBox("Password must be at least six characters")
        ElseIf TextOne.Text = TextTwo.Text Then
            Password(Hash(TextOne.Text))
            Me.Close()
            Me.Dispose()
        End If
    End Sub

    Sub ChangePassword()
        Dim TextOne As TextBox = Me.Controls("TextOne")
        Dim TextTwo As TextBox = Me.Controls("TextTwo")
        Dim TextThree As TextBox = Me.Controls("TextThree")
        If Not Password() = Hash(TextOne.Text) Then
            MsgBox("Current Password is incorrect. Please try again.")
        ElseIf String.IsNullOrEmpty(TextTwo.Text) Then
            MsgBox("Password cannot be left blank.")
        ElseIf Not TextTwo.Text = TextThree.Text Then
            MsgBox("Passwords do not match.")
        ElseIf TextTwo.Text.Length < 6 Then
            MsgBox("Password must be at least six characters")
        ElseIf TextTwo.Text = TextThree.Text Then
            Password(Hash(TextTwo.Text))
            Me.Close()
            Me.Dispose()
        End If
    End Sub

    Function Hash(ByVal sSourceData As String) As String
        Dim tmpSource() As Byte = ASCIIEncoding.ASCII.GetBytes(sSourceData)
        Dim HashByte() As Byte = New MD5CryptoServiceProvider().ComputeHash(tmpSource)
        Hash = ByteArrayToString(HashByte)
    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        Dim i As Integer
        Dim sOutput As New StringBuilder(arrInput.Length)
        For i = 0 To arrInput.Length - 1
            sOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function

    Function Compare(ByVal HashOne() As Byte, ByVal HashTwo() As Byte) As Boolean
        Compare = False
        If HashOne.Length = HashTwo.Length Then
            Dim i As Integer
            Do While (i < HashOne.Length) AndAlso (HashOne(i) = HashTwo(i))
                i += 1
            Loop
            If i = HashOne.Length Then
                Compare = True
            End If
        End If
    End Function

    Private Sub Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If e.CloseReason = CloseReason.UserClosing Then
            If Password().Length = 0 Then
                Dim Password As Form = New Password
                Password.ShowDialog()
            End If
        End If

    End Sub

    Function Password(Optional ByVal NewPassword As String = "") As String
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
            If Not String.IsNullOrEmpty(NewPassword) Then Node.InnerText = NewPassword
            Password = Node.InnerText
            XMLDoc.Save(XML)
        Next
    End Function
End Class