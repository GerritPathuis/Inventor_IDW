Imports System.IO
Imports Inventor


Public Class Form1
    Private iProperties As Object

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Directory create and delete
        Dim path As String = "c:\MyDir"
        Try
            ' Determine whether the directory exists.
            If Directory.Exists(path) Then
                MessageBox.Show("That path exists already.")
                Return
            End If

            ' Try to create the directory.
            Dim di As DirectoryInfo = Directory.CreateDirectory(path)
            MessageBox.Show("The directory was created successfully at " & Directory.GetCreationTime(path))

            ' Delete the directory.
            di.Delete()
            MessageBox.Show("The directory was deleted successfully.")

        Catch ee As Exception
            MessageBox.Show("The process failed ", ee.ToString())
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "idw files |*.idw"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

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

    Sub Set_Custom_Property_VirtualComponent(ByVal oOcc As ComponentOccurrence, ByVal PropName As String, ByVal NewValue As String)

        Dim VirtualDef As VirtualComponentDefinition = TryCast(oOcc.Definition, VirtualComponentDefinition)
        Dim oCustomPropertySet As PropertySet = VirtualDef.PropertySets.Item("Inventor User Defined Properties")

        Dim oProperty As Inventor.Property
        Try
            'set new value
            oProperty = oCustomPropertySet.Item(PropName)
            If CBool(oProperty.Value.ToString) Then
                oProperty.Value = NewValue
            End If
        Catch ex As Exception
            'add property with new value
            oProperty = oCustomPropertySet.Add(NewValue, PropName)
        End Try
    End Sub


    'http://modthemachine.typepad.com/my_weblog/2010/02/accessing-iproperties.html
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim information As System.IO.FileInfo
        information = My.Computer.FileSystem.GetFileInfo("C:\Repositories\Inventor_IDW\VTKE-155000.idw")

        TextBox2.Clear()
        TextBox2.Text &= "Name is " & information.FullName & vbCrLf
        TextBox2.Text &= "Last access time is " & information.LastAccessTime & vbCrLf
        TextBox2.Text &= "The length is " & information.Length
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Read Iproperty

        ' Create Apprentice.
        Dim oApprentice As ApprenticeServerComponent
        oApprentice = New ApprenticeServerComponent

        ' Open a document.
        Dim oDoc As ApprenticeServerDocument
        oDoc = oApprentice.Open("C:\Repositories\Inventor_IDW\Test.ipt")
        MsgBox("Opened: " & oDoc.DisplayName)

        ' Get the PropertySets object. 
        Dim oPropSets As PropertySets
        oPropSets = oDoc.PropertySets

        ' Get the design tracking property set. 
        Dim oPropSet As PropertySet
        oPropSet = oPropSets.Item("Design Tracking Properties")

        ' Get the part number iProperty. 
        ' Dim oPartNumiProp As PropertySet

        'oPartNumiProp = CType(oPropSet.Item("Part Number"), PropertySet)

        ' Display the value. 
        'TextBox3.Text = "The part number is: " & oPartNumiProp.Value
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        SetProperty("Piet")
    End Sub

    'Change the auther name in Iproperties
    Private Sub SetProperty(ByVal author As String)
        Dim mApprenticeserver As ApprenticeServerComponent
        mApprenticeserver = New ApprenticeServerComponent

        Dim oApprenticeDoc As ApprenticeServerDocument
        oApprenticeDoc = mApprenticeServer.Open("C:\Repositories\Inventor_IDW\Test.ipt")

        'Get "Inventor Summary Information" PropertySet
        Dim oPropertySet As PropertySet
        oPropertySet = oApprenticeDoc.PropertySets("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")

        'Get Author property
        Dim oProperty As Inventor.Property = oPropertySet.Item("Author")
        oProperty.Value = author

        oApprenticeDoc.PropertySets.FlushToFile()
        oApprenticeDoc.Close()
    End Sub

End Class

