' xmldoc is form level var
Dim xmlDoc = New Xml.XmlDocument()

'Open Theme File

        Dim f = New OpenFileDialog With {
            .InitialDirectory = ThemePath,
            .Filter = "Themes (*.ask)|*.ask*"
        }
       
        f.ShowDialog()       
        If f.FileName = "" then Exit Sub
        xmlDoc.Load(f.FileName)

        Dim N = xmlDoc.SelectSingleNode("Ableton").ChildNodes(0)

        ' load the names of the theme objects into a list
        ' list only the ones with ARGB values, skip the others
        For Each node As XmlNode In N.ChildNodes
            If node.ChildNodes.Count = 4 Then ListThemeObjects.Items.Add(node.Name)
        Next
        f.Dispose()

'Get Current Object Color on List Item Selection

        LabelColorHeader.Text = ""

        Dim R = "", G = "", B = ""
        Dim N = xmlDoc.SelectSingleNode("Ableton").ChildNodes(0)

        ' display the color value of the selected object, skip alpha
        For Each node As XmlNode In N.ChildNodes
            If node.Name = sender.text Then
                R = Int(node.ChildNodes(0).Attributes(0).Value)
                G = Int(node.ChildNodes(1).Attributes(0).Value)
                B = Int(node.ChildNodes(2).Attributes(0).Value)
                LabelColor.BackColor = Color.FromArgb(R, G, B)
                LabelColorHeader.Text = node.Name
                Exit Sub
            End If
        Next

'Update Selected Object RGB Values

        ' update the RGB values for the selected object, leave Alpha alone
        Dim N = xmlDoc.SelectSingleNode("Ableton").ChildNodes(0)

        For Each node As XmlNode In N.ChildNodes
            If node.Name = ListThemeObjects.Text Then
                node.ChildNodes(0).Attributes(0).Value = LabelColor.BackColor.R.ToString
                node.ChildNodes(1).Attributes(0).Value = LabelColor.BackColor.G.ToString
                node.ChildNodes(2).Attributes(0).Value = LabelColor.BackColor.B.ToString
                Exit Sub
            End If
        Next

'Save Theme File

        Dim xmlStr As String = xmlDoc.innerXML.ToString

        ' must write it this way to use encoding.default or if fails in live
        My.Computer.FileSystem.WriteAllText(fileName, xmlStr, False, System.Text.Encoding.Default)
