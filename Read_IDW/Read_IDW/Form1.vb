Imports System.IO
Imports Inventor


Public Class Form1
    Dim filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Dim filepath2 As String = "C:\Repos\Inventor_IDW\READ_IDW\Part_Copy_update2.ipt"
    Dim filepath3 As String = "c:\MyDir"


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Directory create and delete

        Try
            ' Determine whether the directory exists.
            If Directory.Exists(filepath3) Then
                MessageBox.Show("That path exists already.")
                Return
            End If

            ' Try to create the directory.
            Dim di As DirectoryInfo = Directory.CreateDirectory(filepath3)
            MessageBox.Show("The directory was created successfully at " & Directory.GetCreationTime(filepath3))

            ' Delete the directory.
            di.Delete()
            MessageBox.Show("The directory was deleted successfully.")

        Catch ee As Exception
            MessageBox.Show("The process failed ", ee.ToString())
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog With {
            .InitialDirectory = "c:\",
            .Filter = "Part File (*.ipt)|*.ipt" _
            & "|Assembly File (*.iam)|*.iam" _
            & "|Presentation File (*.ipn)|*.ipn" _
            & "|Drawing File (*.idw)|*.idw" _
            & "|Design element File (*.ide)|*.ide",
            .FilterIndex = 1,                   ' *.ipt files
            .RestoreDirectory = True
        }

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    ' Insert code to read the stream here.
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Sub

    'http://modthemachine.typepad.com/my_weblog/2010/02/accessing-iproperties.html
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim information As System.IO.FileInfo
        information = My.Computer.FileSystem.GetFileInfo(filepath1)

        TextBox2.Clear()
        TextBox2.Text &= "Name= " & information.FullName & vbCrLf
        TextBox2.Text &= "Last access time= " & information.LastAccessTime.ToString & vbCrLf
        TextBox2.Text &= "Last write time= " & information.LastWriteTime.ToString & vbCrLf
        TextBox2.Text &= "Length= " & information.Length & vbCrLf
        TextBox2.Text &= "Creation time= " & information.CreationTime & vbCrLf
        TextBox2.Text &= "Attributes no= " & information.Attributes & vbCrLf
        TextBox2.Text &= "File extension= " & information.Extension & vbCrLf
        TextBox2.Text &= "File exists= " & information.Exists & vbCrLf
        TextBox2.Text &= "Directory= " & information.Directory.ToString & vbCrLf
        TextBox2.Text &= "Directory_name= " & information.DirectoryName.ToString & vbCrLf
        TextBox2.Text &= "Read only= " & information.IsReadOnly.ToString & vbCrLf

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Read Iproperty trough apprentice

        ' Create Apprentice.
        Dim invApprentice As ApprenticeServerComponent
        invApprentice = New ApprenticeServerComponent

        ' Open a document.
        Dim invDoc As ApprenticeServerDocument
        invDoc = invApprentice.Open(filepath1)
        'MessageBox.Show("Opened: " & invDoc.DisplayName)

        '== Get the design tracking property set. 
        Dim invDesigninfo As Inventor.PropertySet
        invDesigninfo = invDoc.PropertySets.Item("Design Tracking Properties")
        Dim invPartNumberProperty As Inventor.Property

        invPartNumberProperty = invDesigninfo.Item("Part Number")
        TextBox3.Text = "The part number is: " & invPartNumberProperty.Value & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Project")
        TextBox3.Text &= "Project: " & invPartNumberProperty.Value & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Checked By")
        TextBox3.Text &= "Checked By: " & invPartNumberProperty.Value & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Description")
        TextBox3.Text &= "Description: " & invPartNumberProperty.Value & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Engineer")
        TextBox3.Text &= "Engineer: " & invPartNumberProperty.Value & vbCrLf

        '== Now switch to "Inventor Summary Information"
        invDesigninfo = invDoc.PropertySets.Item("Inventor Summary Information")
        invPartNumberProperty = invDesigninfo.Item("Author")
        TextBox3.Text &= "Author: " & invPartNumberProperty.Value

        'Close everything
        invDoc = Nothing
        invApprentice.Close()
        invApprentice = Nothing
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Set Iproperty (Inventor does NOT run)
        SetProperty("Piet")
    End Sub

    'Change the auther name in Iproperties (Inventor does NOT run)
    Private Sub SetProperty(ByVal author As String)
        Dim mApprenticeserver As ApprenticeServerComponent
        mApprenticeserver = New ApprenticeServerComponent

        Dim oApprenticeDoc As ApprenticeServerDocument
        oApprenticeDoc = mApprenticeserver.Open(filepath1)

        'Get "Inventor Summary Information" PropertySet
        Dim oPropertySet As PropertySet
        oPropertySet = oApprenticeDoc.PropertySets("Inventor Summary Information")
        'oPropertySet = oApprenticeDoc.PropertySets("Inventor Document Summary Information")
        'oPropertySet = oApprenticeDoc.PropertySets("Design Tracking Properties")
        'oPropertySet = oApprenticeDoc.PropertySets("User Defined Properties")

        'Get Author property
        'Dim oProperty As [Property] = oPropertySet.Item("Title")           'id=2
        'Dim oProperty As [Property] = oPropertySet.Item("Subject")         'id=3
        Dim oProperty As [Property] = oPropertySet.Item("Author")           'id=4
        'Dim oProperty As [Property] = oPropertySet.Item("Keywords")        'id=5
        'Dim oProperty As [Property] = oPropertySet.Item("Comments")        'id=6
        'Dim oProperty As [Property] = oPropertySet.Item("Last Saved By")   'id=8
        'Dim oProperty As [Property] = oPropertySet.Item("Revision Number") 'id=9

        oProperty.Value = author

        oApprenticeDoc.PropertySets.FlushToFile()
        oApprenticeDoc.Close()
        MessageBox.Show("Iproperty Author is changed ")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        SaveToNewFile()
    End Sub

    Public Sub SaveToNewFile()
        Dim apprentice As New ApprenticeServerComponent()


        Try
            'Open part
            Dim appDoc As ApprenticeServerDocument = apprentice.Open(filepath1)

            ' Save the file to a new name
            Dim myFileSaveAs As FileSaveAs = apprentice.FileSaveAs

            myFileSaveAs.AddFileToSave(appDoc, filepath2)

            myFileSaveAs.ExecuteSaveCopyAs()
            appDoc.Close()

        Catch ex As Exception
            Dim attr As FileAttributes = (New FileInfo(filepath1)).Attributes
            MessageBox.Show(String.Format("Error: {0}", ex.Message))
            If (attr And FileAttributes.ReadOnly) > 0 Then
                MessageBox.Show("The file is read-only.")
            End If
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Set Ipropertie (Inventor must run)
        ' Get the Inventor Application object.
        Dim invApp As Inventor.Application
        invApp = CType(GetObject(, "Inventor.Application"), Application)
        ' Get the active document.
        Dim doc As Inventor.Document
        doc = invApp.ActiveDocument
        ' Get the "Design Tracking Properties" property set.
        Dim designTrackPropSet As Inventor.PropertySet
        designTrackPropSet = doc.PropertySets.Item("Design Tracking Properties")

        ' Get the "Description" property from the property set.
        Dim descProp As Inventor.Property
        descProp = designTrackPropSet.Item("Description")
        ' Set the value of the property using the current value of the text box.
        MessageBox.Show(descProp.ToString)
        descProp.Value = TextBox5.Text
        ' Get the "Part Number" property from the property set.

        Dim partNumProp As Inventor.Property
        partNumProp = designTrackPropSet.Item("Part Number")
        ' Set the value of the property using the current value of the text box.
        partNumProp.Value = TextBox4.Text
    End Sub
End Class

