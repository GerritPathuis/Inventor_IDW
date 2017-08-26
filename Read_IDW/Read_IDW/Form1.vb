Imports System.IO
Imports Inventor



Public Class Form1
    Private iProperties As Object

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
        'openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim invApp As Inventor.Application
        invApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")

        Dim Doc As Inventor.Document
        Doc = invApp.ActiveDocument

        UpdateCustomiProperty(Doc, "MYPROPERTY", TextBox1.Text)

        Doc.Update()
        'Set_Custom_Property_VirtualComponent(ByVal oOcc As ComponentOccurrence, "Title", "FAN BOTTOM CASING2")
        Set_Custom_Property_VirtualComponent("oOcc.name", "Title", "FAN BOTTOM CASING2")
    End Sub

    Sub Set_Custom_Property_VirtualComponent(ByVal oOcc As ComponentOccurrence, ByVal PropName As String, ByVal NewValue As String)

        Dim VirtualDef As VirtualComponentDefinition = TryCast(oOcc.Definition, VirtualComponentDefinition)
        Dim oCustomPropertySet As PropertySet = VirtualDef.PropertySets.Item("Inventor User Defined Properties")

        Dim oProperty As Inventor.Property
        Try
            'set new value
            oProperty = oCustomPropertySet.Item(PropName)
            If oProperty.Value.ToString Then
                oProperty.Value = NewValue
            End If
        Catch ex As Exception
            'add property with new value
            oProperty = oCustomPropertySet.Add(NewValue, PropName)
        End Try
    End Sub

    Private Sub UpdateCustomiProperty(ByRef Doc As Inventor.Document, ByRef PropertyName As String, ByRef PropertyValue As Object)
        ' Get the custom property set. 
        Dim customPropSet As Inventor.PropertySet
        customPropSet = Doc.PropertySets.Item("Inventor User Defined Properties")

        ' Get the existing property, if it exists. 
        Dim prop As Inventor.Property = Nothing
        Dim propExists As Boolean = True
        Try
            prop = customPropSet.Item(PropertyName)
        Catch ex As Exception
            propExists = False
        End Try

        ' Check to see if the property was successfully obtained. 
        If Not propExists Then
            ' Failed to get the existing property so create a new one. 
            prop = customPropSet.Add(PropertyValue, PropertyName)
        Else
            ' Change the value of the existing property. 
            prop.Value = PropertyValue
        End If
    End Sub
    'http://modthemachine.typepad.com/my_weblog/2010/02/accessing-iproperties.html
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim information As System.IO.FileInfo
        information = My.Computer.FileSystem.GetFileInfo("C:\Repositories\Inventor_IDW\VTKE-155000.idw")

        MsgBox("The file's full name is " & information.FullName & ".")
        MsgBox("Last access time is " & information.LastAccessTime & ".")
        MsgBox("The length is " & information.Length & ".")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' I have a simple code below trying to iterate through all the assembly component And copy the iProperties > Occurrece Name (i.e. "Part1:1") And paste it to iProperties > Project tab > Part Number field for each individual part in assembly.

        ' set a reference to the assembly component definintion.
        ' This assumes an assembly document is open.
        ''Dim oAsmCompDef As AssemblyComponentDefinition
        ''oAsmCompDef = ThisApplication.ActiveDocument.ComponentDefinition

        '''Iterate through all of the occurrences
        ''Dim oOccurrence As ComponentOccurrence
        ''For Each oOccurrence In oAsmCompDef.Occurrences
        ''    Dim oName As String
        ''    oName = oOccurrence.Name
        ''    iProperties.Value(oOccurrence.Name, "Project", "Part Number") = oName
        ''    MessageBox.Show(oOccurrence.Name, "iLogic")
        ''Next

        ' The Debug message shows the correct output For every iteration, but when I check each part's properties they all have the Occurrence Name of the last part in the aseembly.
        '  I am obivously doing something wrong With the code, but I can't get my head around it at the moment and was hoping someone could help me.
    End Sub
End Class

