Imports System.IO
Imports Inventor

Public Class Form1
    Public filepath1 As String = "C:\Repos\Inventor_IDW\Read_IDW\Part.ipt"
    Public filepath2 As String = "C:\Repos\Inventor_IDW\READ_IDW\Part_update2.ipt"
    Public filepath3 As String = "c:\MyDir"
    Public filepath4 As String = "C:\Temp\Flat_2.dxf"
    Public filepath5 As String = "C:\Inventor_tst\Assembly1.idw"


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
        ' Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog With {
            .InitialDirectory = "c:\Inventor test files\",
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
                filepath1 = openFileDialog1.FileName
                TextBox6.Text = filepath1.ToString
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
            End Try
        End If
    End Sub

    'http://modthemachine.typepad.com/my_weblog/2010/02/accessing-iproperties.html
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim information As System.IO.FileInfo
        MessageBox.Show(filepath1.ToString)
        information = My.Computer.FileSystem.GetFileInfo(filepath1.ToString)

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
        'Read Iproperty via apprentice

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

        'invPartNumberProperty = invDesigninfo.Item("DOC_NUMBER")
        'TextBox3.Text = "The DOC_NUMBER is: " & invPartNumberProperty.Value.ToString & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Part Number")
        TextBox3.Text = "The part number is: " & invPartNumberProperty.Value.ToString & vbCrLf


        invPartNumberProperty = invDesigninfo.Item("Project")
        TextBox3.Text &= "Project: " & invPartNumberProperty.Value.ToString & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Checked By")
        TextBox3.Text &= "Checked By: " & invPartNumberProperty.Value.ToString & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Description")
        TextBox3.Text &= "Description: " & invPartNumberProperty.Value.ToString & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Engineer")
        TextBox3.Text &= "Engineer: " & invPartNumberProperty.Value.ToString & vbCrLf

        invPartNumberProperty = invDesigninfo.Item("Date Checked")
        TextBox3.Text &= "Date Checked: " & invPartNumberProperty.Value.ToString & vbCrLf

        invDesigninfo.Item("Date Checked").Value = Now      'Today modified

        '== Now switch to "Inventor Summary Information"
        invDesigninfo = invDoc.PropertySets.Item("Inventor Summary Information")
        invPartNumberProperty = invDesigninfo.Item("Author")
        TextBox3.Text &= "Author: " & invPartNumberProperty.Value.ToString

        invDoc.PropertySets.FlushToFile()   'Write to file
        invDoc.Close()
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
        'oPropertySet = oApprenticeDoc.PropertySets("Inventor User Defined Properties")

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
        'Dim apprentice As New ApprenticeServerComponent()

        Dim mApprenticeserver As ApprenticeServerComponent
        mApprenticeserver = New ApprenticeServerComponent

        Dim AppDoc As Inventor.ApprenticeServerDocument
        AppDoc = mApprenticeserver.Open(filepath1)


        Try
            ' Save the file to a new name
            Dim myFileSaveAs As FileSaveAs = mApprenticeserver.FileSaveAs   'Loopt hier vast!!

            myFileSaveAs.AddFileToSave(AppDoc, filepath2)

            myFileSaveAs.ExecuteSaveCopyAs()
            appDoc.Close()
        Catch ex As Exception
            Dim attr As FileAttributes = (New FileInfo(filepath1)).Attributes
            MessageBox.Show(String.Format("Error: {0}", ex.Message))
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        If GetInventorApplication() Then
            ' Get the Inventor Application object.
            Dim invApp As Inventor.Application
            invApp = CType(GetObject(, "Inventor.Application"), Application)

            ' Open a document on disk
            Dim doc As Inventor.Document
            doc = invApp.Documents.Open(filepath1, False)

            ' Get the "Design Tracking Properties" property set.
            Dim designTrackPropSet As Inventor.PropertySet
            designTrackPropSet = doc.PropertySets.Item("Design Tracking Properties")

            ' Get the "Description" property from the property set.
            Dim descProp As Inventor.Property
            descProp = designTrackPropSet.Item("Description")

            ' Set the value of the property using the current value of the text box.
            MessageBox.Show(descProp.Value.ToString)
            descProp.Value = TextBox5.Text

            ' Get the "Part Number" property from the property set.
            Dim partNumProp As Inventor.Property
            partNumProp = designTrackPropSet.Item("Part Number")
            ' Set the value of the property using the current value of the text box.
            partNumProp.Value = TextBox4.Text
            doc.Save()
            doc.Close(False)
        Else
            MessageBox.Show("Inventor is nog niet gestart")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        'Inventor moet open zijn, dit werkt.
        If GetInventorApplication() Then

            Dim invDoc As Inventor.Document
            Dim invApp As Inventor.Application
            invApp = CType(GetObject(, "Inventor.Application"), Application)

            ' Open a document on disk
            invDoc = invApp.Documents.Open(filepath1, False)

            'Get "Inventor Summary Information" PropertySet
            Dim invCustomPropertySet As PropertySet
            invCustomPropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties")

            'Declare properties
            Dim strText As String = "fannnn"
            Dim dblValue As Double = 3.14
            Dim dtDate As Date = Now
            Dim blYesOrNo As Boolean = True

            'Create properties
            Dim invProperty As Inventor.Property
            Try
                invProperty = invCustomPropertySet.Add(strText, "Test test")
                invProperty = invCustomPropertySet.Add(dblValue, "Test value")
                invProperty = invCustomPropertySet.Add(dtDate, "Test Date")
                invProperty = invCustomPropertySet.Add(blYesOrNo, "Test Yes or No")
                MessageBox.Show("Iproperty Author is changed ")
            Catch ee As Exception
                MessageBox.Show("Properties already exist..")
            End Try
        Else
            MessageBox.Show("Inventor is nog niet gestart")
        End If
    End Sub

    Private Function GetInventorApplication() As Boolean
        Try
            inventorApplication = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application"), Inventor.Application)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        WriteSheetMetalDXF()
    End Sub
    'This function works !!!!
    Public Sub WriteSheetMetalDXF()
        If GetInventorApplication() Then
            Dim oDoc As Inventor.PartDocument
            Dim invApp As Inventor.Application
            Dim oDataIO As DataIO
            Dim sOut As String

            invApp = CType(GetObject(, "Inventor.Application"), Application)

            ' Get the active document.  This assumes it is a part document.
            oDoc = CType(invApp.Documents.Open(filepath1, False), PartDocument)

            ' Get the DataIO object.
            oDataIO = oDoc.ComponentDefinition.DataIO

            ' Build the string that defines the format of the DXF file.
            sOut = "FLAT PATTERN DXF?AcadVersion=R12&OuterProfileLayer=Outer"

            ' Create the DXF file.
            Try
                oDataIO.WriteDataToFile(sOut, filepath4)
            Catch ee As Exception
                MessageBox.Show("Properties already exist..")
            End Try
            MsgBox(filepath4.ToString & " file created")
        Else
            MessageBox.Show("Inventor is nog niet gestart")
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim invPartDoc As Inventor.PartDocument
        Dim invApp As Inventor.Application

        invApp = CType(GetObject(, "Inventor.Application"), Application)
        invPartDoc = CType(invApp.Documents.Open(filepath1, False), PartDocument)

        TextBox1.Clear()
        ' Get the volume of the part. This will be returned in
        ' cubic centimeters.
        Dim dVolume As Double
        dVolume = invPartDoc.ComponentDefinition.MassProperties.Volume
        TextBox1.Text = "Volume is " & dVolume.ToString & vbCrLf

        ' Get the UnitsOfMeasure object which is used to do unit conversions.
        Dim oUOM As UnitsOfMeasure
        oUOM = invPartDoc.UnitsOfMeasure
        TextBox1.Text &= "Units are " & oUOM.GetStringFromType(oUOM.LengthUnits).ToString & vbCrLf

        '------------------------ 
        ' Convert the volume to the current document units.
        Dim str As String
        str = oUOM.GetStringFromValue(dVolume, oUOM.GetStringFromType(oUOM.LengthUnits) & "^3")
        TextBox1.Text &= "Volume is " & str.ToString & vbCrLf

        '------------------------ 
        ' Get the part number property 
        'see http://modthemachine.typepad.com/my_weblog/2010/02/accessing-iproperties.html
        Dim invP As Inventor.Property
        invP = invPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number")
        TextBox1.Text &= "Part number is " & invP.Value.ToString & vbCrLf

        '------------------------ 
        ' Get the part number property 
        Dim inv As Inventor.Property
        inv = invPartDoc.PropertySets.Item("Design Tracking Properties").Item("Project")
        TextBox1.Text &= "Project is " & inv.Value.ToString & vbCrLf
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TestiPropertyUpdate()
    End Sub
    Private Sub TestiPropertyUpdate()
        'see http://modthemachine.typepad.com/my_weblog/2010/02/custom-iproperties.html
        ' Connect to a running instance of Inventor. 
        ' Watch out for the wrapped line. 
        Dim invApp As Inventor.Application
        invApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application"), Application)
        invApp.SilentOperation = vbTrue

        ' Get the active document. 
        Dim Doc As Inventor.Document
        Doc = CType(invApp.Documents.Open(filepath1, False), Document)

        TextBox1.Clear()
        ' Update or create the custom iProperty. 
        'UpdateCustomiProperty(Doc, "DOC_NUMBER", "D12345")
        ShowCustomiProperty(Doc, "DOC_NUMBER")
        ShowCustomiProperty(Doc, "ITEM_NR")
        ShowCustomiProperty(Doc, "DOC_REV")
        ShowCustomiProperty(Doc, "DOC_STATUS")
        ShowCustomiProperty(Doc, "Thickness")
        ShowCustomiProperty(Doc, "PART_MATERIAL")
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
            MessageBox.Show("Property not found, create new one")
            'prop = customPropSet.Add(PropertyValue, PropertyName)
        Else
            ' Change the value of the existing property. 
            'prop.Value = PropertyValue
            TextBox1.Text &= "Property value is " & prop.Value & vbCrLf
        End If
    End Sub

    Private Sub ShowCustomiProperty(ByRef Doc As Inventor.Document, ByRef PropertyName As String)
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
        If propExists Then
            TextBox1.Text &= "Property value is " & prop.Value & vbCrLf
        End If
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ' Dim myStream As Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog With {
            .InitialDirectory = "c:\Inventor test files\",
            .Filter = "Part File (*.ipt)|*.ipt" _
            & "|Assembly File (*.iam)|*.iam" _
            & "|Presentation File (*.ipn)|*.ipn" _
            & "|Drawing File (*.idw)|*.idw" _
            & "|Design element File (*.ide)|*.ide",
            .FilterIndex = 2,                   ' *.ipt files
            .RestoreDirectory = True
        }

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                filepath1 = openFileDialog1.FileName
                TextBox6.Text = filepath1.ToString
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
            End Try
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Qbom()
    End Sub

    Private Sub Qbom()
        TextBox1.Clear()

        Dim invApp As Inventor.Application
        invApp = CType(GetObject(, "Inventor.Application"), Application)
        invApp.SilentOperation = vbTrue

        Dim oDoc As Inventor.Document
        oDoc = CType(invApp.Documents.Open(filepath1, False), Document)

        Try
            Dim oBOM As Inventor.BOM
            oBOM = oDoc.ComponentDefinition.BOM
            oBOM.StructuredViewFirstLevelOnly = True
            oBOM.StructuredViewEnabled = True

            Dim oBOMView As Inventor.BOMView
            oBOMView = oBOM.BOMViews.Item("Structured")

            '-------------------------
            Dim oRow As BOMRow
            Dim oCompDef As ComponentDefinition
            Dim oPropSet As PropertySet
            Dim i As Integer
            For i = 1 To oBOMView.BOMRows.Count
                oRow = oBOMView.BOMRows.Item(i)
                oCompDef = oRow.ComponentDefinitions.Item(1)
                oPropSet = oCompDef.Document.PropertySets.Item("Design Tracking Properties")

                TextBox1.Text &= "Item= " & oRow.ItemNumber & " Quantity=" & oRow.ItemQuantity
                TextBox1.Text &= " Part= " & oPropSet.Item("Part Number").Value
                TextBox1.Text &= " Desc= " & oPropSet.Item("Description").Value
                TextBox1.Text &= " Stock= " & oPropSet.Item("Stock Number").Value & vbCrLf

                oPropSet = oCompDef.Document.PropertySets.Item("Inventor User Defined Properties")
                TextBox1.Text &= "Doc_nr= " & oPropSet.Item("DOC_NUMBER").Value
                TextBox1.Text &= " Status= " & oPropSet.Item("DOC_STATUS").Value
                TextBox1.Text &= " Doc_rev= " & oPropSet.Item("DOC_REV").Value
                TextBox1.Text &= " Item_nr= " & oPropSet.Item("ITEM_NR").Value & vbCrLf
            Next
        Catch Ex As Exception
            MessageBox.Show("No BOM in this drawing " & Ex.Message)
        Finally

        End Try

    End Sub
    '-----------------------------
    Public Function TitleBlockVersion(ByVal VersionNum As String) As Boolean
        TitleBlockVersion = False

        Dim invApp As Inventor.Application
        invApp = CType(GetObject(, "Inventor.Application"), Application)
        invApp.SilentOperation = vbTrue

        'Dim oDrawDoc As DrawingDocument
        Dim oDoc As Inventor.Document
        oDrawDoc = CType(invApp.Documents.Open(filepath5, False), Document)


        ' Create the new title block defintion.
        Dim oTitleBlockDef As TitleBlockDefinition
        On Error GoTo Errorhandler
        oTitleBlockDef = oDrawDoc.TitleBlockDefinitions.Item("SMALL") ' this is our standard title block

        Dim oSketch As DrawingSketch
        Call oTitleBlockDef.Edit(oSketch)

        Dim Counter As Integer
        Dim Name As String
        Dim VersionName As String
        VersionName = "TITLE BLOCK VERSION: " & VersionNum

        'Loop thru and find the approved by box
        For Counter = 1 To oSketch.TextBoxes.Count
            Name = oSketch.TextBoxes.Item(Counter).Text
            If Name = VersionName Then
                TitleBlockVersion = True
            End If

        Next Counter
        Call oTitleBlockDef.ExitEdit(True)
        Exit Function

Errorhandler:
        TitleBlockVersion = False
        Exit Function

    End Function


    Public Sub TitleBlockCopy()

        'Open Inventor
        Dim oApp As Inventor.Application
        oApp = CType(GetObject(, "Inventor.Application"), Application)
        oApp.SilentOperation = vbTrue

        'Open document
        Dim oDoc As Inventor.Document
        oCurrentDocument = CType(oApp.Documents.Open(filepath5, False), Document)
        MessageBox.Show("Current 1 doc is " & oCurrentDocument.ActiveSheet.TitleBlock.Name)
        'MessageBox.Show("Current 2 doc is " & thisdoc.path)

        'Check to see if the titleblock version is current
        Dim Current As Boolean
        Current = TitleBlockVersion("001")

        ' quit if already current
        If Current Then
            Exit Sub
        End If

        MessageBox.Show("Template Standard.idw is stored at " & oApp.FileOptions.TemplatesPath)
        Dim TemplatePath As String
        TemplatePath = oApp.FileOptions.TemplatesPath & "Standard.idw"




        'Set a reference to the document's active title block name
        Dim TitleBlockName As String
        TitleBlockName = oCurrentDocument.ActiveSheet.TitleBlock.Name

        'Open the template file
        Dim oTemplateDocument As Inventor.Document
        oTemplateDocument = oCurrentDocument.Documents.Open(TemplatePath)

        'oTemplateDocument = CType(oApp.Documents.Open(filepath1, False), Document)

        'Check to see if the template has the same title block as the currentdocument
        Dim RefTitleBlockDef As TitleBlockDefinition
        Dim TitleBlockExists As Boolean
        TitleBlockExists = False
        For Each RefTitleBlockDef In oTemplateDocument.TitleBlockDefinitions
            If RefTitleBlockDef.Name = TitleBlockName Then
                TitleBlockExists = True
                Debug.Print("Found Title Block")
            End If
        Next

        If Not TitleBlockExists Then
            TitleBlockName = "SMALL" ' this is our default title block
        End If

        ' Get the new source title block definition.
        Dim oSourceTitleBlockDef As TitleBlockDefinition
        oSourceTitleBlockDef = oTemplateDocument.TitleBlockDefinitions.Item(TitleBlockName)

        'Wipe out any references to the existing title block
        Dim oSheet As Sheet
        oCurrentDocument.Activate()

        For Each oSheet In oCurrentDocument.Sheets
            oSheet.Activate()
            On Error Resume Next
            oSheet.TitleBlock.Delete()
        Next

        'Delete the existing titleblock
        On Error Resume Next
        oCurrentDocument.TitleBlockDefinitions.Item(TitleBlockName).Delete()

        'Copy the Template Title Block to the current file
        Dim oNewTitleBlockDef As TitleBlockDefinition
        oNewTitleBlockDef = oSourceTitleBlockDef.CopyTo(oCurrentDocument)

        oTemplateDocument.Close()

        ' Iterate through the sheets.
        For Each oSheet In oCurrentDocument.Sheets
            oSheet.Activate()
            Call oSheet.AddTitleBlock(oNewTitleBlockDef)
        Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TitleBlockCopy()
    End Sub
End Class

