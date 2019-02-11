Imports System.Xml
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

''' <summary>
''' Class to print Tabular data from a database
''' </summary>
''' <remarks>This class prints out data retrieved from a database, where the data can be grouped
''' by specified columns, etc.
''' Hopefully, we can use this with various projects
''' </remarks>
'''


Public Class TblPrn
    ' PrintConfig: ArrayList containing groups with print Data
    Private PrintConfig As Hashtable ' The formatting definition
    Private reader As XmlReader
    Private InConfig As Boolean = False ' Flag that we've encountered the <config> element
    ' grpAttribList: Hashtable keyed by element name, value= Array of legal attrib names
    Private grpAttribList As Hashtable = New Hashtable()
    Private grpLvls As ArrayList = New ArrayList()

    ' This block of variables are those used by the PrintDocument functions
    Private prnDoc As PrintDocument
    Private allData As ArrayList
    Private yPos As Single
    Private pgWdth As Single
    Private pgHeight As Single
    Private yMax As Single
    Private curRow As Integer
    Private rowHeight As Integer = 20
    Private pageIndexes As ArrayList = New ArrayList()
    Dim inPass2 As Boolean
    Dim curPage As Integer
    Private BoxLvl As Integer
    'Private tableprintdtd As String

    Private Enum GRP_TYPE As Integer
        GpTyGrp
        GpTyHdr
    End Enum

    Public Sub New()
        InitGrpAttList()            ' Initialize grpAttribList
        prnDoc = New PrintDocument()

        ' Set up some Default Page Settings
        With prnDoc.DefaultPageSettings.Margins
            .Left = 50.0
            .Top = 50.0
            .Right = 50.0
            .Bottom = 50.0
        End With
    End Sub


    ''' <summary>
    ''' Initialize standard arrayw
    ''' </summary>
    ''' <remarks></remarks>

    Private Sub InitGrpAttList()
        Dim cellAtt() As String = {"percent", "font", "align", "attribs", "src", "align"}
        Dim grpAtt() As String = {"grpsrc", "split_grp", "header", "footer", "pointsabove", "pointsbelow",
                                    "lindent", "rindent", "font", "grpsrc", "pointsabove", "pointsbelow",
                                    "outerbox"}
        Dim dfltAtt() As String = {"font", "lindent", "rindent", "attribs"}

        With grpAttribList
            .Add("cell", cellAtt)
            .Add("group", grpAtt)
            .Add("defaultcell", dfltAtt)
            .Add("header", {"font", "align", "pointsabove", "pointsbelow"})
            .Add("body", {"font", "align", "src", "pointsabove", "pointsbelow", "rowspacing"})
            .Add("pageheader", {"font", "align", "pointsabove", "pointsbelow"})
            .Add("src", {"type", "value"})
        End With
    End Sub

    Private Sub ShowPrintConfig()
        Dim lvl As Hashtable = New Hashtable
        Dim msg As String = ""
        Dim indent As String = ""

        lvl = PrintConfig

        While lvl.Contains("subgroup")
            msg = String.Concat(indent, msg, lvl("name"), vbCrLf)

            If lvl.Count > 1 Then
                For Each k As String In lvl.Keys
                    If Not k.Equals("subgroup") Then
                        msg = String.Concat(indent, msg, String.Format("   key={{{0}}} value='{1}'", k, lvl(k)), vbCrLf)
                    End If
                Next

                lvl = lvl("subgroup")
            End If

            indent = String.Concat(indent, "   ")
        End While

        msg = String.Concat(indent, msg, lvl("name"))

        If lvl.Count > 0 Then
            For Each k As String In lvl.Keys
                msg = String.Concat(vbCrLf, indent, msg, String.Format("   key={{{0}}} value='{1}'", k, lvl(k)))
            Next
        End If

        MessageBox.Show(msg, "Your Config")
    End Sub

    ''' <summary>
    ''' Set up print config from an XML file
    ''' </summary>
    ''' <param name="fn">The name of the file to read</param>
    ''' <remarks>Reads in the printout specifications from a file formatted with the XML setupo
    '''     for the printout
    ''' </remarks>

    Public Function Config_from_file(fn As String)
        Dim rslt As Boolean
        'create_DTD("TblPrint.dtd")

        Using fs = New FileStream(fn, IO.FileMode.Open)
            Dim settings As New XmlReaderSettings()
            'settings.DtdProcessing = DtdProcessing.Parse
            'settings.ValidationType = ValidationType.DTD
            reader = XmlReader.Create(fs, settings)
            rslt = BuildConfigFromReader()
            reader.Close()
            fs.Close()
        End Using

        Return rslt
    End Function

    ''' <summary>
    ''' Reads in the XML printout specifications from a string
    ''' </summary>
    ''' <param name="xmlString">The string specifying the printout format</param>
    ''' <remarks>
    '''     Reads in the specification from a string, which is set up with the XML definition
    '''     for the printout.
    ''' </remarks>

    Public Function Config_from_string(xmlString As String)
        Dim rslt As Boolean
        ' Create the XmlReader object.
        Dim settings As New XmlReaderSettings()
        'settings.DtdProcessing = DtdProcessing.Parse
        'settings.ValidationType = ValidationType.DTD
        'create_DTD("TblPrint.dtd")

        reader = XmlReader.Create(New StringReader(xmlString), settings)
        rslt = BuildConfigFromReader()
        reader.Close()
        Return rslt
    End Function

    ''' <summary>
    ''' Sub to fill the config data using the reader created by the config_from_* () sub
    ''' </summary>
    ''' <remarks>
    '''    This sub uses the global XmlReader "reader" created in the calling sub
    ''' </remarks>

    Private Function BuildConfigFromReader() As Boolean
        Dim prevGrp As Hashtable = Nothing    ' The previous (parent) group

        If IsNothing(PrintConfig) Then
            PrintConfig = New Hashtable()
        Else
            If PrintConfig.Count > 0 Then
                PrintConfig.Clear()
            End If
        End If

        ' Set up a default for the <defaultcell> <font>
        PrintConfig.Add("defaultcell", New Hashtable())
        CType(PrintConfig("defaultcell"), Hashtable).Add("font", New Font("Arial", 10))
        CType(PrintConfig("defaultcell"), Hashtable).Add("name", "defaultcell")

        Do Until reader.NodeType.Equals(XmlNodeType.Element)
            reader.Read()
        Loop

        If Not reader.Name.Equals("config") Then
            MessageBox.Show("The first element must be <config>, not " & reader.Name)
            Return False
        End If

        PrintConfig.Add("parent", Nothing)
        PrintConfig.Add("name", reader.Name)
        prevGrp = PrintConfig
        grpLvls.Add(PrintConfig)
        InConfig = True     ' Flag that we've hit the <config> Element

        ' Now continue

        While reader.Read()
            Select Case reader.NodeType
                Case XmlNodeType.Element
                    ' It may be possible to eliminate this check, perhaps with using a DTD file
                    If reader.Name.Equals("config") Then
                        MessageBox.Show("Element <config> can only be the top-level Element")
                        If InConfig = False Then
                            ' Initialize the PrintConfig Hashtable
                            PrintConfig = New Hashtable()
                        Else    ' Error- Throw Exception?
                        End If

                        Exit Select
                    Else            ' Anything but <config>
                        If reader.Name.Equals("font") Then
                            Select Case prevGrp("name")
                                Case "pageheader", "header", "defaultcell", "body", "cell", "src", "dochead"
                                    ' No statements - these are the valid parents
                                Case Else
                                    MessageBox.Show(String.Format("<{0}> cannot be a parent for <font>",
                                                        prevGrp("name")))
                                    Return Nothing
                            End Select

                            Dim newFont As Font = AddFont(prevGrp("name"))

                            If Not IsNothing(newFont) Then
                                prevGrp("font") = (newFont)
                            End If
                            Continue While
                        Else
                            If InConfig = False Then
                                MessageBox.Show("Must be inside the <config> Element", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                                ' Throw exception?
                            End If
                        End If
                    End If      ' End if reader.Name.Equals("config")

                    Dim ng As Object = AddGroup(prevGrp)
                    If IsNothing(ng) Then
                        Return False
                    End If

                    If Not IsNothing(ng) Then
                        ' We don't need to add a grpLvl if the element is of the form "<.... />"
                        If Not reader.IsEmptyElement Then
                            prevGrp = ng
                            grpLvls.Add(ng)
                        End If
                    End If
                Case XmlNodeType.EndElement
                    If (grpLvls.Count) > reader.Depth Then
                        grpLvls.RemoveAt(grpLvls.Count - 1)

                        If grpLvls.Count > 0 Then
                            prevGrp = grpLvls(grpLvls.Count - 1)
                        End If
                    End If
                Case XmlNodeType.Text
                    ' Any other types simply ignored
            End Select
        End While

        'ShowPrintConfig()       ' For Debugging ...
        Return True
    End Function

    ' ******************************************************************************
    ' SetGrpAttribs() - Reads the Attibutes for an Element and stores them into a
    '       newly-created HashTable, keyed by the Attribute Name
    '
    ' Passed:   (1) attribs - Array containing list of permissible attributes
    ' Returns:  The newly-created HashTable (possibly empty) containing all specified attribs
    ' ******************************************************************************
    ''' <summary>
    ''' setGrpAttribs() - Reads the Attributes for an Element and stores them into a
    ''' Hashtable (initialized here in this function).
    ''' </summary>
    ''' <returns>The newly-created Hashtable (possibly empty, containing all
    ''' specified Attributes.</returns>
    ''' <remarks>Each Item in the Hashtable is keyed by the Attribute name and the
    ''' value is the value from the XML table.
    ''' </remarks>
    ''' 
    Private Function SetGrpAttribs() As Hashtable
        Dim cellSpec = New Hashtable()

        If reader.HasAttributes() Then
            Dim myAttribs = grpAttribList(reader.Name)      ' For convenience
            Dim oldname As String = reader.Name

            'For a As Integer = 0 To myAttribs.GetUpperBound(0)
            'For a As Integer = 0 To reader.AttributeCount
            While reader.MoveToNextAttribute()

                Dim attr As String = StrConv(reader.Name, VbStrConv.Lowercase)
                Dim attrVal As String = reader.GetAttribute(attr)

                If Not IsNothing(attrVal) Then
                    Select Case attr
                        Case "align"
                            Select Case StrConv(attrVal, VbStrConv.Lowercase)
                                Case "r", "right"
                                    cellSpec(attr) = StringAlignment.Far
                                Case "c", "center"
                                    cellSpec(attr) = StringAlignment.Center
                                Case Else
                                    cellSpec(attr) = StringAlignment.Near
                            End Select
                        Case "underline", "bold", "italic", "strikethrough"
                            MessageBox.Show("We have Font Styles being checked in the ""SetGrpAttribs()"" function")
                            If Not cellSpec.Contains("style") Then
                                cellSpec("style") = New FontStyle()
                                cellSpec("style") = FontStyle.Regular
                            End If

                            Select Case attr
                                Case "underline"
                                    If attrVal = "yes" Then
                                        cellSpec("style") = cellSpec("style") Or FontStyle.Underline
                                    End If
                                Case "bold"
                                    If attrVal = "yes" Then
                                        cellSpec("style") = cellSpec("style") Or FontStyle.Bold
                                    End If
                                Case "italic"
                                    If attrVal = "yes" Then
                                        cellSpec("style") = cellSpec("style") Or FontStyle.Italic
                                    End If
                                Case "underline"
                                    If attrVal = "yes" Then
                                        cellSpec("style") = cellSpec("style") Or FontStyle.Underline
                                    End If
                            End Select
                        Case "boxed"
                            cellSpec("boxed") = attrVal
                        Case Else
                            cellSpec(attr) = reader.GetAttribute(attr)
                    End Select
                End If
                'Next
            End While

            reader.MoveToElement()      ' Return to begin of element
            Dim newname As String = reader.Name
        End If

        Return cellSpec
    End Function

    ''' <summary>
    ''' Create a new Font
    ''' </summary>
    ''' <returns>The new font if valid parameters are provided, else Nothing</returns>
    ''' <remarks></remarks>

    Private Function AddFont(parentName As String) As Font
        Dim fontAttribs() As String = {"name", "size", "style"}
        Dim fam As String = CType(PrintConfig("defaultcell")("font"), Font).Name
        Dim size As Single = CType(PrintConfig("defaultcell")("font"), Font).Size
        Dim style As FontStyle = CType(PrintConfig("defaultcell")("font"), Font).Style


        ' TODO: Perhaps add features to create a font with default parameters???
        If reader.HasAttributes Then
            Dim styleVal As FontStyle

            For a As Integer = 0 To reader.AttributeCount - 1
                reader.MoveToAttribute(a)

                Select Case reader.Name.ToLower
                    Case "name"
                        fam = reader.Value
                    Case "size"
                        size = CType(reader.Value, Single)
                    Case "underline", "bold", "strikethrough", "italic"
                        Select Case reader.Name
                            Case "underline"
                                styleVal = FontStyle.Underline
                            Case "bold"
                                styleVal = FontStyle.Bold
                            Case "strikethrough"
                                styleVal = FontStyle.Strikeout
                            Case "italic"
                                styleVal = FontStyle.Italic
                        End Select

                        If reader.Value.Equals("yes") Then
                            style = style Or styleVal
                        ElseIf reader.Value.Equals("no") Then
                            style = style And (Not styleVal)
                        End If
                End Select
            Next
        End If

        If (Not IsNothing(fam)) And (Not IsNothing(size)) Then
            Return New Font(fam, size, style)
        End If

        Return Nothing
    End Function

    Private Function ElementVerify(child As String, parent As String) As Boolean
        Select Case parent
            Case "font", "src"
                MessageBox.Show(String.Format("<{0}>:  <{1}> has no children", child, parent), "ERROR!")
                Return False
        End Select

        Select Case child
            Case "config"
                If reader.Depth = 1 Then
                    Return True
                Else
                    MessageBox.Show(String.Format("<config> must be top-level, not child of <{0}>", parent))
                    Return False
                End If
            Case "pageheader", "dochead", "defaultcell"
                If parent.Equals("config") Then
                    Return True
                End If
            Case "group", "body"
                Dim p As String() = {"config", "group"}
                If p.Contains(parent) Then
                    Return True
                End If
            Case "cell"
                Dim p As String() = {"pageheader", "dochead", "group", "body", "header", "subheader"}
                If p.Contains(parent) Then
                    Return True
                End If
            Case "src"
                If parent.Equals("cell") Then
                    Return True
                End If
            Case "header", "subheader", "footer"
                Dim p As String() = {"group", "body"}
                If p.Contains(parent) Then
                    Return True
                End If
            Case "font"
                ' <font> is available almost everywhere
                Dim p As String() = {"config"}
                If Not p.Contains(parent) Then
                    Return True
                End If
            Case "boxed"
                Dim p As String() = {"group", "body"}
                If p.Contains(parent) Then
                    Return True
                End If
            Case Else
                MessageBox.Show(String.Format("<{0}> is not a valid element", child), "ERROR!")
                Return False
        End Select

        MessageBox.Show(String.Format("<{0}> is not a child of <{1}>", child, parent), "ERROR!")
        Return False
    End Function

    ' ******************************************************************************
    ' AddGroup() - Add a group or  to the list of the elements in the config
    '       This includes all types : <group>, <header>, <footer>, etc
    '       On entry, the reader is positioned at the StartElement
    ' ******************************************************************************
    ''' <summary>
    ''' addGroup() - Add a new group to the list of Elements in the config
    ''' </summary>
    ''' <param name="parentGrp">The group definition Hashtable for the parent of this Element
    ''' </param>
    ''' <returns>
    ''' If the new group is a true subgroup, the new subgroup,
    ''' else Nothing in case of &lt;cell&gt; or anything that MUST be a top-level item
    ''' </returns>
    ''' <remarks>
    ''' This function sets up a definition to a new group.
    ''' It returns this new Hashtable so that the caller can recognize it, and probably
    ''' reset the pointer to the "previous" group to be this
    ''' </remarks>

    Private Function AddGroup(parentGrp As Hashtable) As Hashtable
        If Not ElementVerify(reader.Name, parentGrp("name")) Then
            Return Nothing
        End If
        Dim newGrp As Hashtable = Nothing

        Select Case reader.Name
            Case "defaultcell", "pageheader"
                If Not reader.Depth = 1 Then
                    Dim msg As String = "<{0}> must be Top-Level element (below <config>)"
                    MessageBox.Show(String.Format(msg, reader.Name))
                End If
        End Select

        Select Case reader.Name
            Case "group", "body", "header", "subheader", "dochead", "src"
                newGrp = SetGrpAttribs()
                newGrp("name") = reader.Name
                newGrp("parent") = parentGrp

                Select Case newGrp("name")
                    Case "group", "body"
                        parentGrp("subgroup") = newGrp
                    Case Else
                        parentGrp(newGrp("name")) = newGrp
                End Select
            Case "font"
            Case "cell"
                Select Case parentGrp("name")
                    Case "pageheader", "header", "subheader", "defaultcell", "body", "src", "dochead"
                    Case Else
                        MessageBox.Show(String.Format("<{0}> cannot be a parent for <cell>",
                                                        parentGrp("name")))
                End Select

                If Not parentGrp.ContainsKey("cells") Then
                    parentGrp.Add("cells", New ArrayList())
                End If

                newGrp = SetGrpAttribs()
                newGrp("name") = reader.Name
                CType(parentGrp.Item("cells"), ArrayList).Add(newGrp)
            'newGrp = Nothing    ' We don't need to return the cell def

            ' <defaultcell> already exists, so we create a temporary hashtable, then copy
            ' valid keys
            Case "defaultcell"
                Dim tmpCell As Hashtable = New Hashtable()

                If tmpCell.ContainsKey("align") Then
                    CType(PrintConfig("defaultcell"), Hashtable).Add("align", tmpCell("align"))
                End If

                tmpCell = Nothing
                newGrp = PrintConfig("defaultcell")
            Case "pageheader"
                newGrp = SetGrpAttribs()
                newGrp("name") = reader.Name
                parentGrp(reader.Name) = newGrp
                newGrp("parent") = parentGrp
                'newGrp = Nothing    ' We don't want to change the parentGroup for these
        End Select

        Return newGrp
    End Function

    ''' <summary>
    ''' Render_report() - The entry pointy rendering the entire report
    ''' </summary>
    ''' <param name="dataRows">
    ''' The ArrayList containing the rows of the Query result
    ''' </param>
    ''' <remarks>The entry point for rendering the entire report.  dataRows is an ArrayList.
    ''' Each Object in the ArrayList is a Hashtable keyed by the column names
    ''' </remarks>

    Public Sub Render_report(dataRows As ArrayList)
        allData = dataRows
        AddHandler prnDoc.PrintPage, AddressOf Me.PrnDoc_PrintPage

        Dim pd As New PrintDialog With {
            .Document = prnDoc,
            .AllowSomePages = True
        }
        Dim rslt As DialogResult = pd.ShowDialog()

        If rslt = DialogResult.OK Then
            curRow = 0      ' Begin with first row
            curPage = 1
            inPass2 = False

            If pageIndexes.Count Then
                pageIndexes.Clear()
            End If

            prnDoc.Print()

        End If
    End Sub

    Private Sub Process_page(ev As PrintPageEventArgs)
        Dim curLevel As Hashtable = PrintConfig
        Dim pgMaxIdx As Integer

        pgWdth = ev.MarginBounds.Right - ev.MarginBounds.Left
        pgHeight = ev.MarginBounds.Bottom - ev.MarginBounds.Top
        ' The following is an arbitrary setting to allow full printing of the bottom line
        yMax = ev.MarginBounds.Bottom - 10

        yPos = ev.MarginBounds.Top

        ' Print PageHeader
        If curLevel.Contains("pageheader") Then
            ' TODO: Set up Attributes
            yPos += Render_row_cells(curLevel("pageheader"), ev)
        End If

        If curRow = 0 Then
            If curLevel.Contains("dochead") Then
                yPos += Render_row_cells(curLevel("dochead"), ev)
            End If
        End If

        If inPass2 Then
            pgMaxIdx = pageIndexes(0)
            pageIndexes.RemoveAt(0)
        Else
            pgMaxIdx = allData.Count
        End If
        ' If First Page and a <docheader> is needed...

        ' Now begin parsing down through the groupings
        If curLevel.Contains("subgroup") Then
            curLevel = curLevel("subgroup")

            If curLevel("name").Equals("group") Then
                Process_group(curLevel, pgMaxIdx, ev)
            Else
                'ElseIf curLevel("name") = "body" Then
                Process_body(curLevel, pgMaxIdx, ev)
                'Else
                '    MessageBox.Show("""" & curLevel("name") & """ is not a valid printing group")
                '    ev.HasMorePages = False
                '    Return
            End If
        End If
    End Sub

    ''' <summary>
    ''' Event Handler for the PrintPage Event for the PrintDocument class
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ev"></param>
    ''' <remarks>Handles the printing of a single page.  it prints the document header
    ''' (if it's the first page), the page header, keeps up with what's printable on the first
    ''' page, etc
    ''' </remarks>

    Private Sub PrnDoc_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)

        ' On first page, do a dry run through the entire dataset to determine the total
        ' pages and fine-tune the pagination
        If curRow = 0 Then
            curPage = 1
            inPass2 = False

            While curRow < allData.Count
                Process_page(ev)
                pageIndexes.Add(curRow)
                curPage += 1
            End While

            ' Now Set up to do the actual printing
            inPass2 = True
            curPage = 1
            curRow = 0
        End If

        Process_page(ev)
        curPage += 1


        If curRow < allData.Count Then
            If curPage < 20 Then
                ev.HasMorePages = True
            Else
                ev.HasMorePages = False
            End If
        Else            ' Else at end of data
            ev.HasMorePages = False
        End If
    End Sub

    Private Sub TpDoBox(myTop As Integer, boxType As String, ev As PrintPageEventArgs)
        Dim myPen As New Pen(Color.Black)

        yPos += 5     ' Allow a little more space at the bottom

        ' We draw one box in any case
        Select Case boxType
            Case "heavy"
            Case "double"
                myPen.Width = 3
            Case "light"
                myPen.Width = 1
        End Select

        ev.Graphics.DrawRectangle(myPen, ev.MarginBounds.Left, myTop, pgWdth, yPos - myTop)

        'For "double", draw a light box inside the outer box
        If boxType = "double" Then
            myPen.Width = 1
            ev.Graphics.DrawRectangle(myPen, ev.MarginBounds.Left + 4, myTop + 4, pgWdth - 8, yPos - myTop - 8)
        End If

        yPos += 5    ' Add some space below line
    End Sub

    ''' <summary>
    ''' Process a group
    ''' </summary>
    ''' <param name="myLvl">The Hashtable defining this group</param>
    ''' <param name="maxRow">The maximum index + 1 (stopping-point) for this group</param>
    ''' <param name="ev">The PrintPageEventArgs for this PrintDocument</param>
    ''' <remarks>This is the main processing routine for the printout.  The PrintPage handler comes
    '''   here unless there are no groups but simply a body.  A group's headers and are processed here,
    '''   and then it recursuvely calls itself for a subgroup until a &lt;body&gt; is encountered.
    ''' </remarks>

    Private Sub Process_group(myLvl As Hashtable, maxRow As Integer, ev As PrintPageEventArgs)
        ' TODO: process attributes
        Dim myRow As Hashtable = allData(curRow)
        Dim grpMax As Integer = curRow
        Dim grpTop As Integer = yPos

        If (myLvl.Contains("boxed")) Then
            BoxLvl += 1
            yPos += 10
        End If

        While curRow < maxRow
            Dim curName As String = allData(curRow)(myLvl("grpsrc"))

            ' Render the main header
            If myLvl.Contains("header") Then
                yPos += Render_row_cells(myLvl("header"), ev)
            End If

            ' Render the Subheader if applicable
            If myLvl.Contains("subheader") Then
                yPos += Render_row_cells(myLvl("subheader"), ev)
            End If

            ' Find index of last row fitting into this group
            While (grpMax < allData.Count)
                If Not allData(grpMax)(myLvl("grpsrc")).Equals(curName) Then
                    Exit While
                End If

                grpMax += 1
            End While

            If myLvl.Contains("subgroup") Then
                If CType(myLvl("subgroup"), Hashtable)("name").Equals("group") Then

                    Process_group(myLvl("subgroup"), grpMax, ev)

                    If (yPos >= yMax) Or (curRow >= allData.Count) Then
                        Exit Sub
                    End If
                Else    ' It must be <body>
                    Process_body(myLvl("subgroup"), grpMax, ev)

                    If (yPos >= yMax) Or (curRow >= allData.Count) Then
                        If (myLvl.Contains("boxed")) Then
                            TpDoBox(grpTop, myLvl("boxed"), ev)

                            If BoxLvl > 0 Then
                                BoxLvl -= 1
                            End If
                        End If

                        Exit Sub
                    End If
                End If
            End If

            If (myLvl.Contains("boxed")) Then
                TpDoBox(grpTop, myLvl("boxed"), ev)
            End If

            If myLvl.Contains("splitgroup") Then
                yPos = yMax + 1
            End If
        End While

        If BoxLvl > 0 Then
            BoxLvl -= 1
        End If
    End Sub

    ''' <summary>
    ''' Handles printing of the body (actual data)
    ''' </summary>
    ''' <param name="myLvl">The hashtable defining the current group-level</param>
    ''' <param name="maxRow">The index (+1) of the last row of data to print in this grouping</param>
    ''' <param name="ev">The PrintPageEventArgs for the PrintDocument</param>
    ''' <remarks></remarks>

    Private Sub Process_body(myLvl As Hashtable, maxRow As Integer, ev As PrintPageEventArgs)
        Dim cellWidths As ArrayList = CalcCellAryWidths(myLvl("cells"), ev.MarginBounds.Width)

        ' Unless we have only a single row to print, don't print a single row by itself on the bottom line.
        ' In this case, bump yPos up past the limit and return
        If (maxRow - curRow > 1) Then
            If (yMax - yPos < 20) Then
                yPos = yMax + 1
                Exit Sub
            End If
        End If

        ' Render the main header
        If myLvl.Contains("header") Then
            yPos += Render_row_cells(myLvl("header"), ev)
        End If


        For myRow As Integer = curRow To maxRow - 1
            ' TODO: process attributes (if necessary)

            If yPos >= yMax Then
                ' TODO: Any cleanup to do?
                Exit Sub
            End If

            yPos += Render_row_cells(myLvl, ev)
            curRow += 1
        Next
    End Sub

    ''' <summary>
    ''' Converts the "percent" attributes in an Array of &lt;cell&gt;'s to an ArrayList of widths
    ''' </summary>
    ''' <param name="cA">The Array of &lt;cell&gt;s</param>
    ''' <param name="totWidth">The total with of the printable line</param>
    ''' <returns>The ArrayList containing the widths</returns>
    ''' <remarks></remarks>

    Private Function CalcCellAryWidths(cA As ArrayList, totWidth As Single) As ArrayList
        Dim widthAry As New ArrayList()

        For c As Integer = 0 To cA.Count - 1
            widthAry.Add(CType(cA(c)("percent"), Single) * totWidth / 100)
        Next

        Return widthAry
    End Function

    ''' <summary>
    ''' Render a row of cells
    ''' </summary>
    ''' <param name="curgrp">The Hashtable defining this current group</param>
    ''' <param name="ev">The PrintPageEventArgs</param>
    ''' <returns>The total units to add to the yPos</returns>
    ''' <remarks>This function returns all the units to be added to yPos and
    ''' the caller is responsible for updating this value.  This function could
    ''' handle all this itself but we'll leave it like this as it's possible that
    ''' some future modification might not want yPos updated for some reason,</remarks>

    Private Function Render_row_cells(curgrp As Hashtable, ev As PrintPageEventArgs) As Single
        Dim maxHeight As Single = 0
        Dim addHeight As Single = 0
        Dim rowTop As Single = yPos
        ' TODO: Set up attributes

        If curgrp("pointsabove") Then
            rowTop += CType(curgrp("pointsabove"), Single)
        End If

        If curgrp.Contains("lineabove") Then
            ev.Graphics.DrawLine(Pens.Black, ev.MarginBounds.Left + (BoxLvl * 10), yPos,
                                     ev.MarginBounds.Right - (BoxLvl * 20), yPos)
            addHeight += 2
            rowTop += 2
        End If

        If curgrp.Contains("cells") Then
            ' Set Width=0 because we add the previous cell width to the X value before drawing
            Dim cellRect As New RectangleF(ev.MarginBounds.Left + (BoxLvl * 10), rowTop, 0,
                                               ev.MarginBounds.Bottom - rowTop)
            Dim cellWdths(curgrp("cells").Count - 1)
            ' Copy the "cells" arraylist to an Array to simplify matters
            Dim cellAry(CType(curgrp("cells"), ArrayList).Count - 1) As Hashtable
            CType(curgrp("cells"), ArrayList).CopyTo(cellAry)

            ' Calculate Cell Width
            For c As Integer = 0 To cellAry.GetUpperBound(0)
                Dim fnt As Font = Nothing

                If curgrp.Contains("font") Then
                    fnt = curgrp("font")
                End If

                cellRect.X += cellRect.Width
                cellRect.Width = CType(cellAry(c)("percent"), Single) *
                    (ev.MarginBounds.Right - (BoxLvl * 20) - ev.MarginBounds.Left) / 100
                Dim cH As Single = RenderCell(curgrp("name"), cellAry(c), cellRect, fnt, ev)

                If cH > maxHeight Then
                    maxHeight = cH
                End If

            Next
        End If

        If curgrp.Contains("linebelow") Then
            ev.Graphics.DrawLine(Pens.Black, ev.MarginBounds.Left + (BoxLvl * 10), yPos + maxHeight + addHeight,
                                     ev.MarginBounds.Right - (BoxLvl * 20), yPos + maxHeight + addHeight)
            addHeight += 4
        End If

        If curgrp.Contains("pointsabove") Then
            maxHeight += curgrp("pointsabove")
        End If

        If curgrp.Contains("pointsbelow") Then
            maxHeight += CType(curgrp("pointsbelow"), Single)
        End If

        Return maxHeight + addHeight
    End Function

    Private Function RenderCell(grpTyp As String, cellAry As Hashtable,
                                     ByVal cellRect As RectangleF, parntFont As Font, ev As PrintPageEventArgs) As Single
        ' TODO set up attributes
        Dim fnt As Font
        Dim str As String = ""

        ' Determine text
        ' TODO: We may eliminate this top-level Select, it seems that we
        ' now don't need to check for this
        Select Case grpTyp
            Case "group"
                If cellAry.Contains("groupsource") Then
                    Select Case cellAry("groupsource")

                    End Select
                End If
            Case "header", "body", "subheader", "dochead", "pageheader"
                If cellAry.Contains("src") Then
                    Dim r As Hashtable = cellAry("src")

                    Select Case r("type")
                        Case "data"
                            Dim d As Hashtable = allData(curRow)

                            If IsDBNull(d(r("value"))) Then
                                str = ""
                            Else
                                str = d(r("value"))
                            End If
                        Case "text"
                            str = r("value")
                        Case Else
                            str = "<N/A>"
                    End Select
                End If
        End Select

        If cellAry.Contains("font") Then
            fnt = cellAry("font")
        ElseIf Not IsNothing(parntFont) Then
            fnt = parntFont
        ElseIf PrintConfig.Contains("defaultcell") And CType(PrintConfig("defaultcell"),
                                Hashtable).Contains("font") Then
            fnt = PrintConfig("defaultcell")("font")
        Else
            fnt = New Font("Arial", 8.0F)
        End If

        If inPass2 Then             ' We don't actually draw the data till Pass 2
            If cellAry.Contains("lindent") Then
                cellRect.X += cellAry("lindent")
                cellRect.Width -= cellAry("lindent")
            End If

            If cellAry.Contains("align") Then
                Dim sFmt As StringFormat = New StringFormat With {
                    .Alignment = cellAry("align")
                }
                ev.Graphics.DrawString(str, fnt, New SolidBrush(Color.Black), cellRect, sFmt)
                sFmt.Dispose()
            Else
                ev.Graphics.DrawString(str, fnt, New SolidBrush(Color.Black), cellRect)
            End If
        End If

        Return ev.Graphics.MeasureString(str, fnt, cellRect.Width).Height

    End Function

    ''' <summary>
    ''' Returns the PrintDocument for the class
    ''' </summary>
    ''' <returns>prnDoc - the PrintDocument for the class</returns>
    ''' <remarks></remarks>

    Public ReadOnly Property PrntDoc As PrintDocument
        Get
            Return Me.prnDoc
        End Get
    End Property

    ' ******************************************************************************
    ' AddDefaults() - Sets the defaults for the config
    ' ******************************************************************************

    Private Sub AddDefaults()
        Dim attribs() As String = {"font", "lindent", "rindent", "attribs"}
    End Sub

    ''' <summary>
    ''' Add a Page Header to the list of the elements in the config
    ''' </summary>

    Private Sub AddPagehdr()
        Dim attribs() As String = {}
    End Sub

End Class
