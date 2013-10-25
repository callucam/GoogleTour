#Region "Imports directives"

Imports System.Reflection
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

#End Region


Public Class Form1

    Dim XPlaceMark As XElement
    Dim XModel As XElement


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing
        'Dim NameSheet As Excel.Worksheet = Nothing
        'Dim LookAtSheet As Excel.Worksheet = Nothing
        'Dim PlaceMarkSheet As Excel.Worksheet = Nothing

        Dim oCells As Excel.Range = Nothing
        oXL = New Excel.Application
        oXL.Visible = False
        oWBs = oXL.Workbooks
        'oWB = oWBs.Open("C:\Resource Documents\Resources\Simulation\Google Earth Animation\SetupViewFromLocation.xlsm")
        'NameSheet = oWB.Worksheets(1)
        'LookAtSheet = oWB.Worksheets(2)
        'PlaceMarkSheet = oWB.Worksheets(3)

        oWB = oWBs.Open("C:\Resource Documents\Resources\Simulation\Google Earth Animation\Template.xlsx")
        Dim Document As Excel.Worksheet = oWB.Worksheets(1)
        'Dim LookAtSheet1 As Excel.Worksheet = oWB.Worksheets(2)
        'Dim Style1 As Excel.Worksheet = oWB.Worksheets(3)
        Dim Folder As Excel.Worksheet = oWB.Worksheets(2)
        'Dim LookAt2 As Excel.Worksheet = oWB.Worksheets(5)
        'Dim Style2 As Excel.Worksheet = oWB.Worksheets(6)
        Dim PlaceMark As Excel.Worksheet = oWB.Worksheets(3)
        Dim Tour As Excel.Worksheet = oWB.Worksheets(4)
        Dim FlyTo As Excel.Worksheet = oWB.Worksheets(5)

        'Tour.Range("a1").Value = TourName.Text

        Dim m2 As Date = ns1when.Value
        Dim m3 As Date = CDate(Date.FromOADate(CDbl(ns1when.Value.ToOADate()) + 1 / 60 / 60 / 24))
        Dim RowCount As Integer = 100

        Dim horizFov As String = ns2horizFov.Text
        Dim longitude As Double = ns1longitude.Text
        Dim longitudeMax As Double = ns1longitudeMax.Text
        Dim latitude As Double = ns1latitude.Text
        Dim latitudeMax As Double = ns1latitudeMax.Text
        Dim altitude As Double = ns1altitude.Text
        Dim altitudeMax As Double = ns1altitudeMax.Text
        Dim heading As Double = ns1heading.Text
        Dim headingMax As Double = ns1headingMax.Text
        Dim tilt As Double = ns1tilt.Text
        Dim tiltMax As Double = ns1tiltMax.Text
        Dim range As Double = ns1range.Text
        Dim rangeMax As Double = ns1rangeMax.Text
        Dim altitudeMode As String = ns2altitudeMode.Text
        Dim duration As Double = ns2duration.Text
        Dim flyToMode As String = ns2flyToMode.Text

        For i = 0 To RowCount
            FlyTo.Range("a2").Offset(i, 0).Value = altitudeMode
            FlyTo.Range("b2").Offset(i, 0).Value = horizFov
            FlyTo.Range("c2").Offset(i, 0).Value = Year(m2) & "-" & Format(Month(m2), "00") & "-" & Format(Day(m2), "00") & "T" & Format(Hour(m2), "00") & ":" & Format(Minute(m2), "00") & ":" & Format(Second(m2), "00") & "Z"
            FlyTo.Range("d2").Offset(i, 0).Value = Year(m2) & "-" & Format(Month(m2), "00") & "-" & Format(Day(m2), "00") & "T" & Format(Hour(m2), "00") & ":" & Format(Minute(m2), "00") & ":" & Format(Second(m2), "00") & "Z"
            FlyTo.Range("e2").Offset(i, 0).Value = longitude + (longitudeMax - longitude) / RowCount * i
            FlyTo.Range("f2").Offset(i, 0).Value = latitude + (latitudeMax - latitude) / RowCount * i
            FlyTo.Range("g2").Offset(i, 0).Value = altitude + (altitudeMax - altitude) / RowCount * i
            FlyTo.Range("h2").Offset(i, 0).Value = heading + (headingMax - heading) / RowCount * i
            FlyTo.Range("i2").Offset(i, 0).Value = tilt + (tiltMax - tilt) / RowCount * i
            FlyTo.Range("j2").Offset(i, 0).Value = range + (rangeMax - range) / RowCount * i
            FlyTo.Range("k2").Offset(i, 0).Value = duration
            FlyTo.Range("l2").Offset(i, 0).Value = flyToMode
        Next


        Dim MyPath As String = "C:\Resource Documents\Resources\Simulation\Google Earth Animation\"
        Dim MyFile As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & ".xml"
        Dim NewName As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & ".kml"

        Dim zz As Excel.XmlMap = oWB.XmlMaps("kml_Map")

        oWB.SaveAsXMLData(MyPath & MyFile, zz)

        oWB.Close(False)

        If Dir(MyPath & MyFile) <> "" Then
            My.Computer.FileSystem.RenameFile(MyPath & MyFile, NewName)
            'Name MyPath & MyFile As MyPath & NewName
        Else
            MsgBox("File not found")
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim initialDirectory As String = "C:\Resource Documents\Resources\Simulation\Google Earth Animation\"
        OpenFileDialog1.InitialDirectory = initialDirectory
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.DefaultExt = "kml"
        OpenFileDialog1.AddExtension = True
        OpenFileDialog1.Filter = "KML files (*.kml)|*.kml|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()

        Dim k As XNamespace = "http://www.opengis.net/kml/2.2"

        PlaceMarkName.Text = OpenFileDialog1.FileName

        XPlaceMark = XElement.Load(OpenFileDialog1.FileName)

        Dim coordinates As String = (XPlaceMark.Elements(k + "Document").Elements(k + "Placemark").Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)

        Dim firstcomma As Integer = (InStr(coordinates, ","))

        Dim longitude As Double = Microsoft.VisualBasic.Left(coordinates, firstcomma - 1)

        coordinates = Microsoft.VisualBasic.Right(coordinates, Len(coordinates) - firstcomma)

        firstcomma = (InStr(coordinates, ","))

        Dim lattitude As Double = Microsoft.VisualBasic.Left(coordinates, firstcomma - 1)

        Dim altitude As Double = Microsoft.VisualBasic.Right(coordinates, Len(coordinates) - firstcomma)

        ns1longitude.Text = longitude
        ns1latitude.Text = lattitude
        ns1altitude.Text = altitude

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim initialDirectory As String = "C:\Resource Documents\Resources\Simulation\Google Earth Animation\"
        OpenFileDialog2.InitialDirectory = initialDirectory
        OpenFileDialog2.FileName = ""
        OpenFileDialog2.DefaultExt = "kml"
        OpenFileDialog2.AddExtension = True
        OpenFileDialog2.Filter = "KML files (*.kml)|*.kml|All files (*.*)|*.*"
        OpenFileDialog2.ShowDialog()

        Dim k As XNamespace = "http://www.opengis.net/kml/2.2"

        ModelName.Text = OpenFileDialog2.FileName

        XModel = XElement.Load(OpenFileDialog2.FileName)

    End Sub
End Class
