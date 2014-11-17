'Add view from stationary location
'Add direct export to XML
'Add CCPE logo
'Add parameter call out -- speed, heading, heel, trim, draft

'Models: 

'crane barge
'sea trials
'turning cycle
'seakeeping to buoy


#Region "Imports directives"

Imports System.Reflection
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

#End Region


Public Class Form1

    Dim XPlaceMark(10) As XElement
    Dim XPlaceMark1 As XElement
    Dim XPlaceMark2 As XElement
    Dim XModel As XElement
    Dim DaeName(8) As String
    Dim DaeNameSteps As Integer
    Dim pi = 3.14159265358979
    Dim EarthRadius = 6378.1 * 1000
    Dim NPlacemarks As Integer
    Dim k As XNamespace = "http://www.opengis.net/kml/2.2"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim mystring As String = ""

        ' Set view properties

        Dim horizFov As String = ns2horizFov.Text
        Dim heading As Double = ns1heading.Text
        Dim headingMax As Double = ns1headingMax.Text
        Dim tilt As Double = ns1tilt.Text
        Dim tiltMax As Double = ns1tiltMax.Text
        Dim range As Double = ns1range.Text
        Dim rangeMax As Double = ns1rangeMax.Text

        ' Set model properties

        Dim PMheading As Double = SpeedMin.Text
        Dim PMheadingMax1 As Double = SpeedMax.Text
        Dim PMtilt As Double = PMrollMin.Text
        Dim PMtiltMax1 As Double = PMrollMax.Text
        Dim PMroll As Double = PMpitchMin.Text
        Dim PMrollMax1 As Double = PMpitchMax.Text
        Dim altitudeMode As String = ns2altitudeMode.Text
        Dim duration As Double = ns2duration.Text 'this is a percentage of the total length of tour, to the total duration.
        Dim flyToMode As String = ns2flyToMode.Text

        ' Load Placemarks

        Dim LonLatAlt
        Dim longitudes(10) As Double
        Dim latitudes(10) As Double
        Dim altitudes(10) As Double

        If FromLatLonRadioButton.Checked = True Then
            LonLatAlt = LoadPlacemarks(PmFolderListBox.Items(0))
        Else
            LonLatAlt = LoadPlacemarks(PmReferenceTextBox.Text)
        End If

        longitudes = LonLatAlt(0)
        latitudes = LonLatAlt(1)
        altitudes = LonLatAlt(2)

        LoadModel()

        ' Write to Excel

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing
        Dim oWB1 As Excel.Workbook = Nothing
        Dim oWB2 As Excel.Workbook = Nothing
        Dim oWB3 As Excel.Workbook = Nothing

        oXL = New Excel.Application
        oXL.Visible = False
        oWBs = oXL.Workbooks

        oWB = oWBs.Open("C:\Google Earth Tour\Template of Model.xlsx")
        Dim Document As Excel.Worksheet = oWB.Worksheets(1)

        Dim Folder As Excel.Worksheet = oWB.Worksheets(2)

        Dim PlaceMark As Excel.Worksheet = oWB.Worksheets(3)
        Dim Tour As Excel.Worksheet = oWB.Worksheets(4)
        Dim FlyTo As Excel.Worksheet = oWB.Worksheets(5)

        oWB1 = oWBs.Open("C:\Google Earth Tour\Template for Track.xlsx")

        oWB3 = oWBs.Open("C:\Google Earth Tour\Template for Placemark Data.xlsx")

        Dim PlaceMark1 As Excel.Worksheet = oWB1.Worksheets(1)
        Dim When1 As Excel.Worksheet = oWB1.Worksheets(2)
        Dim Coord1 As Excel.Worksheet = oWB1.Worksheets(3)

        Dim Sheet1 As Excel.Worksheet

        Dim PMDataHeader As Excel.Worksheet = oWB3.Worksheets(1)
        Dim PMDataTable As Excel.Worksheet = oWB3.Worksheets(2)

        Dim m2 As Date = ns1when.Value
        Dim m3 As Date = m2

        Dim DistanceArray(10) As Double
        Dim GlobalBearingArray(10) As Double
        Dim XArray(10) As Double
        Dim YArray(10) As Double
        Dim LocalBearingArray(10) As Double
        Dim OrientationArray(10) As Double
        Dim SpeedMinText As Double = SpeedMin.Text
        Dim SpeedMaxText As Double = SpeedMax.Text
        Dim SpeedArray(10) As Double
        Dim TimeArray(10) As Integer
        Dim VxArray(10) As Double
        Dim VyArray(10) As Double
        Dim AxArray(10) As Double
        Dim BxArray(10) As Double
        Dim AyArray(10) As Double
        Dim ByArray(10) As Double
        Dim j As Integer = 0
        Dim xPosition As Double
        Dim yPosition As Double
        Dim DistanceBetweenXY As Double = 0
        Dim BearingBetweenXY As Double = 0
        Dim OutputLatDeg As Double
        Dim OutputLongDeg As Double

        If FromLatLonRadioButton.Checked = True Then
            For speedindex = 0 To NPlacemarks
                SpeedArray(speedindex) = SpeedMinText + (SpeedMaxText - SpeedMinText) / NPlacemarks * speedindex
            Next
            DistanceArray = DistBetweenPlacemarks(latitudes, longitudes)
            GlobalBearingArray = GlobalBearingBetweenPlacemarks(latitudes, longitudes)
            XArray = xarrayfromdistbearing(DistanceArray, GlobalBearingArray)
            YArray = yarrayfromdistbearing(DistanceArray, GlobalBearingArray)
            LocalBearingArray = LocalBearingBetweenPlacemarks(XArray, YArray)
            OutputLatDeg = latitudes(0) * 180 / pi
            OutputLongDeg = longitudes(0) * 180 / pi
            xPosition = 0
            yPosition = 0
        Else
            oWB2 = oWBs.Open(ExcelSeriesTextBox.Text)

            Sheet1 = oWB2.Worksheets(1)
            NPlacemarks = 10
            XArray(0) = Sheet1.Range("a2").Offset(0, 0).Value
            YArray(0) = Sheet1.Range("b2").Offset(0, 0).Value
            For n = 1 To NPlacemarks
                For speedindex = 0 To NPlacemarks
                    SpeedArray(speedindex) = SpeedMinText + (SpeedMaxText - SpeedMinText) / NPlacemarks * speedindex
                Next
                XArray(n) = Sheet1.Range("a2").Offset(n, 0).Value ' xarrayfromdistbearing(DistanceArray, GlobalBearingArray)
                YArray(n) = Sheet1.Range("b2").Offset(n, 0).Value
                DistanceArray(n) = ((XArray(n) - XArray(0)) ^ 2 + (YArray(n) - YArray(0)) ^ 2) ^ 0.5
                GlobalBearingArray(n) = Math.Atan2(YArray(n) - YArray(0), XArray(n) - XArray(0))
                LocalBearingArray(n) = Math.Atan2(YArray(n) - YArray(n - 1), XArray(n) - XArray(n - 1))

            Next

            xPosition = XArray(0)
            xPosition = YArray(0)

            DistanceBetweenXY = (XArray(0) ^ 2 + YArray(0) ^ 2) ^ 0.5
            BearingBetweenXY = Math.Atan2(yPosition, xPosition) - 90 * pi / 180

            'MsgBox(DistanceBetweenXY & " " & BearingBetweenXY)

            OutputLatDeg = (Math.Asin(Math.Sin(latitudes(0)) * Math.Cos(DistanceBetweenXY / 1000 / 6378.1) + Math.Cos(latitudes(0)) * Math.Sin(DistanceBetweenXY / 1000 / 6378.1) * Math.Cos(BearingBetweenXY))) * 180 / pi
            OutputLongDeg = (longitudes(0) + Math.Atan2(Math.Cos(DistanceBetweenXY / EarthRadius) - Math.Sin(latitudes(0)) * Math.Sin(OutputLatDeg * pi / 180), Math.Sin(BearingBetweenXY) * Math.Sin(DistanceBetweenXY / EarthRadius) * Math.Cos(latitudes(0)))) * 180 / pi - 90

            oWB2.Close(False)
        End If

        For p = 0 To LocalBearingArray.Length - 2
            OrientationArray(p) = LocalBearingArray(p) / 2 + LocalBearingArray(p + 1) / 2
        Next

        OrientationArray(0) = LocalBearingArray(1)

        OrientationArray(NPlacemarks) = LocalBearingArray(NPlacemarks)

        TimeArray = TimeArrayfromDistanceArray(DistanceArray, SpeedArray)
        VxArray = VxArrayfromLocalBearingAndSpeed(OrientationArray, SpeedArray)
        VyArray = VyArrayfromLocalBearingAndSpeed(OrientationArray, SpeedArray)
        AxArray = AArrayfromPositionTimeSpeed(XArray, TimeArray, VxArray)
        BxArray = BArrayfromPositionTimeSpeed(XArray, TimeArray, VxArray)
        AyArray = AArrayfromPositionTimeSpeed(YArray, TimeArray, VyArray)
        ByArray = BArrayfromPositionTimeSpeed(YArray, TimeArray, VyArray)

        Dim ModelX As Double
        Dim ModelY As Double
        Dim ModelBearing As Double

        Dim OutputLatDegPrevious As Double
        Dim OutputLongDegPrevious As Double
        'Dim ModelBearingPrevious As Double

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = TimeArray(NPlacemarks)

        ProgressBar1.Visible = True
        ProgressBar1.Value = ProgressBar1.Minimum

        Dim i As Double = 0
        Dim index As Integer = 0
        Dim TimeIncrementText As Integer = TimeIncrement.Text

        Dim HeadingString As String
        Dim SpeedString As String
        Dim HeelString As String
        Dim TrimString As String
        Dim DraftString As String

        For h = 1 To NPlacemarks

            While i < TimeArray(h)

                xPosition = 1 / 6 * AxArray(h) * (i - TimeArray(h - 1)) ^ 3 + 1 / 2 * BxArray(h) * (i - TimeArray(h - 1)) ^ 2 + VxArray(h - 1) * (i - TimeArray(h - 1)) + XArray(h - 1) '+XArray(0) 
                yPosition = 1 / 6 * AyArray(h) * (i - TimeArray(h - 1)) ^ 3 + 1 / 2 * ByArray(h) * (i - TimeArray(h - 1)) ^ 2 + VyArray(h - 1) * (i - TimeArray(h - 1)) + YArray(h - 1) '+ YArray(0)
                DistanceBetweenXY = (xPosition ^ 2 + yPosition ^ 2) ^ 0.5

                BearingBetweenXY = Math.Atan2(yPosition, xPosition) - 90 * pi / 180
                OutputLatDegPrevious = OutputLatDeg
                OutputLatDeg = (Math.Asin(Math.Sin(latitudes(0)) * Math.Cos(DistanceBetweenXY / 1000 / 6378.1) + Math.Cos(latitudes(0)) * Math.Sin(DistanceBetweenXY / 1000 / 6378.1) * Math.Cos(BearingBetweenXY))) * 180 / pi
                OutputLongDegPrevious = OutputLongDeg
                OutputLongDeg = (longitudes(0) + Math.Atan2(Math.Cos(DistanceBetweenXY / EarthRadius) - Math.Sin(latitudes(0)) * Math.Sin(OutputLatDeg * pi / 180), Math.Sin(BearingBetweenXY) * Math.Sin(DistanceBetweenXY / EarthRadius) * Math.Cos(latitudes(0)))) * 180 / pi - 90
                ProgressBar1.Value = i

                'Set the view

                FlyTo.Range("a2").Offset(index, 0).Value = altitudeMode
                FlyTo.Range("b2").Offset(index, 0).Value = horizFov
                FlyTo.Range("c2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                FlyTo.Range("d2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                When1.Range("a2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                FlyTo.Range("e2").Offset(index, 0).Value = OutputLongDeg 'longitude + (longitudeMax - longitude) / RowCount1 * i
                FlyTo.Range("f2").Offset(index, 0).Value = OutputLatDeg 'latitude + (latitudeMax - latitude) / RowCount1 * i
                FlyTo.Range("g2").Offset(index, 0).Value = altitudes(0) '+ (altitudeMax - altitude) / RowCount1 * i

                Coord1.Range("a2").Offset(index, 0).Value = OutputLongDeg & " " & OutputLatDeg & " " & altitudes(0)

                If LinearHeadingOption.Checked = True Then
                    FlyTo.Range("h2").Offset(index, 0).Value = heading + (headingMax - heading) / TimeArray(NPlacemarks) * i
                Else
                    FlyTo.Range("h2").Offset(index, 0).Value = i Mod 360
                End If
                FlyTo.Range("i2").Offset(index, 0).Value = tilt + (tiltMax - tilt) / TimeArray(NPlacemarks) * i
                FlyTo.Range("j2").Offset(index, 0).Value = range + (rangeMax - range) / TimeArray(NPlacemarks) * i
                FlyTo.Range("k2").Offset(index, 0).Value = duration
                FlyTo.Range("l2").Offset(index, 0).Value = flyToMode



                'Set the model

                PlaceMark.Range("b2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"

                PMDataTable.Range("f2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"


                m3 = CDate(Date.FromOADate(CDbl(m3.ToOADate()) + TimeIncrementText / 60 / 60 / 24))

                PlaceMark.Range("c2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                PMDataTable.Range("g2").Offset(index, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"

                PlaceMark.Range("e2").Offset(index, 0).Value = altitudeMode
                PlaceMark.Range("f2").Offset(index, 0).Value = OutputLongDeg 'longitude + (longitudeMax - longitude) / RowCount1 * i
                PlaceMark.Range("g2").Offset(index, 0).Value = OutputLatDeg 'latitude + (latitudeMax - latitude) / RowCount1 * i
                PlaceMark.Range("h2").Offset(index, 0).Value = altitudes(0) '+ (altitudeMax - altitude) / RowCount1 * i

                PMDataTable.Range("j2").Offset(index, 0).Value = OutputLongDeg & "," & OutputLatDeg & "," & altitudes(0)

                ModelY = Math.Sin((OutputLongDeg - OutputLongDegPrevious) * pi / 180) * Math.Cos(OutputLatDeg * pi / 180)
                ModelX = Math.Cos(OutputLatDegPrevious * pi / 180) * Math.Sin(OutputLatDeg * pi / 180) - Math.Sin(OutputLatDegPrevious * pi / 180) * Math.Cos(OutputLatDeg * pi / 180) * Math.Cos((OutputLongDeg - OutputLongDegPrevious) * pi / 180)
                ModelBearing = Math.Atan2(ModelY, ModelX) * 180 / pi - 90

                PlaceMark.Range("i2").Offset(index, 0).Value = ModelBearing 'PMheading '+ (PMheadingMax1 - PMheading) / RowCount1 * i

                If LinearRollOption.Checked = True Then
                    PlaceMark.Range("j2").Offset(index, 0).Value = PMtilt + (PMtiltMax1 - PMtilt) / TimeArray(NPlacemarks) * i
                    TrimString = "Trim: " & Math.Round(PlaceMark.Range("j2").Offset(index, 0).Value, 1) & "°; "
                Else
                    PlaceMark.Range("j2").Offset(index, 0).Value = RollMagnitude.Text * Math.Sin(2 * pi / RollPeriod.Text * i + RollPhase.Text * pi / 180)
                    TrimString = "Trim: " & Math.Round(PlaceMark.Range("j2").Offset(index, 0).Value, 1) & "°; "
                End If


                If LinearPitchOption.Checked = True Then
                    PlaceMark.Range("k2").Offset(index, 0).Value = PMroll + (PMrollMax1 - PMroll) / TimeArray(NPlacemarks) * i
                    HeelString = "Heel: " & Math.Round(PlaceMark.Range("k2").Offset(index, 0).Value, 1) & "°; "
                Else
                    PlaceMark.Range("k2").Offset(index, 0).Value = PitchMagnitude.Text * Math.Sin(2 * pi / PitchPeriod.Text * i + PitchPhase.Text * pi / 180)
                    HeelString = "Heel: " & Math.Round(PlaceMark.Range("k2").Offset(index, 0).Value, 1) & "°; "
                End If

                PlaceMark.Range("o2").Offset(index, 0).Value = DaeName(j)
                If j = DaeNameSteps Then j = 0 Else j = j + 1

                HeadingString = "Heading: " & Math.Round(ModelBearing + 90, 1) & "°; "
                SpeedString = "Speed: " & Math.Round(DistanceBetweenXY / TimeIncrementText, 1) & "m/s; "
                DraftString = "Draft: " & Math.Round(altitudes(0), 1) & "m ; "




                PMDataTable.Range("b2").Offset(index, 0).Value = HeadingString & SpeedString & DraftString & TrimString & HeelString

                PMDataTable.Range("h2").Offset(index, 0).Value = "#Style_5"

                i = i + TimeIncrementText
                index = index + 1

            End While
        Next

        'Next

        'Dim MyPath As String = "C:\Resource Documents\Resources\Simulation\Google Earth Animation\"
        Dim MyPath As String = "C:\Google Earth Tour\"
        Dim MyFile As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "model.xml"
        Dim NewName As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "model.kml"

        Dim zz As Excel.XmlMap = oWB.XmlMaps("kml_Map")

        oWB.SaveAsXMLData(MyPath & MyFile, zz)

        oWB.Close(False)

        If Dir(MyPath & MyFile) <> "" Then
            My.Computer.FileSystem.RenameFile(MyPath & MyFile, NewName)
            'Name MyPath & MyFile As MyPath & NewName
        Else
            MsgBox("File not found")
        End If

        Dim MyFile1 As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "track.xml"
        Dim NewName1 As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "track.kml"

        Dim yy As Excel.XmlMap = oWB1.XmlMaps("kml_Map")

        oWB1.SaveAsXMLData(MyPath & MyFile1, yy)

        oWB1.Close(False)

        If Dir(MyPath & MyFile1) <> "" Then
            My.Computer.FileSystem.RenameFile(MyPath & MyFile1, NewName1)
            'Name MyPath & MyFile As MyPath & NewName
        Else
            MsgBox("File not found")
        End If


        Dim MyFile2 As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "data.xml"
        Dim NewName2 As String = TourName.Text & " " & Year(m2) & Format(Month(m2), "00") & Format(Day(m2), "00") & " " & Format(Hour(m2), "00") & Format(Minute(m2), "00") & Format(Second(m2), "00") & "data.kml"

        Dim xx As Excel.XmlMap = oWB3.XmlMaps("kml_Map")

        oWB3.SaveAsXMLData(MyPath & MyFile2, xx)

        oWB3.Close(False)

        If Dir(MyPath & MyFile2) <> "" Then
            My.Computer.FileSystem.RenameFile(MyPath & MyFile2, NewName2)
            'Name MyPath & MyFile As MyPath & NewName
        Else
            MsgBox("File not found")
        End If


        ProgressBar1.Value = ProgressBar1.Minimum

    End Sub
    Private Function DistBetweenPlacemarks(latitudes As Double(), longitudes As Double()) As Double()
        Dim lat1 As Double
        Dim lat2 As Double
        Dim lon1 As Double
        Dim lon2 As Double
        Dim ArrayHolder(10) As Double
        For g = 0 To NPlacemarks
            lat1 = latitudes(0)
            lat2 = latitudes(g)
            lon1 = longitudes(0)
            lon2 = longitudes(g)
            ArrayHolder(g) = Math.Acos(Math.Sin(lat1) * Math.Sin(lat2) + Math.Cos(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)) * EarthRadius
        Next
        Return ArrayHolder
    End Function

    Private Function GlobalBearingBetweenPlacemarks(latitudes As Double(), longitudes As Double()) As Double()
        Dim lat1 As Double
        Dim lat2 As Double
        Dim lon1 As Double
        Dim lon2 As Double
        Dim y As Double
        Dim x As Double
        Dim ArrayHolder(10) As Double
        For g = 1 To NPlacemarks
            lat1 = latitudes(0)
            lat2 = latitudes(g)
            lon1 = longitudes(0)
            lon2 = longitudes(g)
            y = Math.Sin(lon2 - lon1) * Math.Cos(lat2)
            x = Math.Cos(lat1) * Math.Sin(lat2) - Math.Sin(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)
            ArrayHolder(g) = Math.Atan2(y, x)
        Next
        Return ArrayHolder
    End Function
    Private Function LocalBearingBetweenPlacemarks(XArray As Double(), YArray As Double()) As Double()
        Dim x1 As Double
        Dim x2 As Double
        Dim y1 As Double
        Dim y2 As Double
        Dim y As Double
        Dim x As Double
        Dim ArrayHolder(10) As Double
        For g = 1 To NPlacemarks
            x1 = XArray(g - 1)
            x2 = XArray(g)
            y1 = YArray(g - 1)
            y2 = YArray(g)
            y = (y2 - y1)
            x = (x2 - x1)
            ArrayHolder(g) = Math.Atan2(y, x)
        Next
        ArrayHolder(0) = ArrayHolder(1)
        Return ArrayHolder
    End Function
    Private Function xarrayfromdistbearing(DistanceArray As Double(), BearingArray As Double()) As Double()
        Dim ArrayHolder(10) As Double
        For g = 0 To NPlacemarks
            ArrayHolder(g) = Math.Sin(BearingArray(g)) * DistanceArray(g)
        Next
        Return ArrayHolder
    End Function
    Private Function yarrayfromdistbearing(DistanceArray As Double(), BearingArray As Double()) As Double()
        Dim ArrayHolder(10) As Double
        For g = 0 To NPlacemarks
            ArrayHolder(g) = Math.Cos(BearingArray(g)) * DistanceArray(g)
        Next
        Return ArrayHolder
    End Function

    Private Function TimeArrayfromDistanceArray(DistanceArray As Double(), SpeedArray As Double()) As Integer()
        Dim ArrayHolder(10) As Integer
        ArrayHolder(0) = DistanceArray(0) * TimeFactor.Text 'DistanceArray(0) / SpeedArray(0)

        For g = 1 To NPlacemarks
            ArrayHolder(g) = DistanceArray(g) * TimeFactor.Text + ArrayHolder(g - 1) 'DistanceArray(k) / SpeedArray(k) + ArrayHolder(k - 1)
        Next

        Return ArrayHolder
    End Function
    Private Function VxArrayfromLocalBearingAndSpeed(LocalBearingArray As Double(), SpeedArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For g = 0 To NPlacemarks
            ArrayHolder(g) = SpeedArray(g) * Math.Cos(LocalBearingArray(g))
            'MsgBox(ArrayHolder(k))
        Next
        Return ArrayHolder
    End Function
    Private Function VyArrayfromLocalBearingAndSpeed(LocalBearingArray As Double(), SpeedArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For g = 0 To NPlacemarks
            ArrayHolder(g) = SpeedArray(g) * Math.Sin(LocalBearingArray(g))
        Next
        Return ArrayHolder
    End Function
    Private Function AArrayfromPositionTimeSpeed(XYArray As Double(), TimeArray As Integer(), VxyArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For g = 1 To NPlacemarks
            ArrayHolder(g) = 6 * ((VxyArray(g) + VxyArray(g - 1)) * (TimeArray(g) - TimeArray(g - 1)) - 2 * (XYArray(g) - XYArray(g - 1))) / (TimeArray(g) - TimeArray(g - 1)) ^ 3
        Next
        Return ArrayHolder
    End Function
    Private Function BArrayfromPositionTimeSpeed(XYArray As Double(), TimeArray As Integer(), VxyArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For g = 1 To NPlacemarks
            ArrayHolder(g) = -2 * ((VxyArray(g) + 2 * VxyArray(g - 1)) * (TimeArray(g) - TimeArray(g - 1)) - 3 * (XYArray(g) - XYArray(g - 1))) / (TimeArray(g) - TimeArray(g - 1)) ^ 2

        Next
        Return ArrayHolder
    End Function
    Private Sub CreateRadiusReference_Click(sender As Object, e As EventArgs) Handles CreateRadiusReference.Click

        NPlacemarks = 0

        Dim LonLatAlt(3) As Double

        LonLatAlt = LoadPlacemarks(RadiusCenter.Text)

        MsgBox(LonLatAlt(0))

        ' NOT COMPLETE

    End Sub
    Private Function LoadPlacemarks(p1 As String) As Object

        'For pm = 0 To NPlacemarks
        '    XPlaceMark(pm) = XElement.Load(p1)
        'Next

        XPlaceMark(0) = XElement.Load(p1)

        Dim coordinates(10) As String

        If (XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark").Count) > 0 Then
            NPlacemarks = XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark").Count - 1
            For pm = 0 To NPlacemarks
                coordinates(pm) = (XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(pm).Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)
                'MsgBox(coordinates(pm))
            Next
        Else
            NPlacemarks = 0
            coordinates(0) = (XPlaceMark(0).Elements(k + "Document").Elements(k + "Placemark").Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)

        End If

        Dim firstcomma(10) As Integer

        For pm = 0 To NPlacemarks

            firstcomma(pm) = (InStr(coordinates(pm), ","))

        Next

        Dim longitudes(10) As Double

        For pm = 0 To NPlacemarks
            longitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        For pm = 0 To NPlacemarks
            coordinates(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        For pm = 0 To NPlacemarks
            firstcomma(pm) = (InStr(coordinates(pm), ","))
        Next

        Dim latitudes(10) As Double

        For pm = 0 To NPlacemarks
            latitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        Dim altitudes(10) As Double

        For pm = 0 To NPlacemarks
            altitudes(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        Return {longitudes, latitudes, altitudes}

    End Function
    Private Sub LoadModel()

        DaeNameSteps = -1

        For h = 0 To DaeModelListBox.Items.Count - 1
            DaeName(h) = DaeModelListBox.Items(h)
            DaeNameSteps = DaeNameSteps + 1
        Next
    End Sub
    Private MouseIsDown As Boolean = False
    Private Sub PmFolderListBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PmFolderListBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub PmFolderListBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PmFolderListBox.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                PmFolderListBox.Items.Add(MyFiles(i))
            Next
        End If
    End Sub
    Private Sub DaeModelListBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DaeModelListBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub DaeModelListBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles DaeModelListBox.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                DaeModelListBox.Items.Add(MyFiles(i))
            Next
        End If
    End Sub
    Private Sub PmReferenceTextBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PmReferenceTextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub PmReferenceTextBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles PmReferenceTextBox.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                PmReferenceTextBox.Text = MyFiles(i)
            Next
        End If
    End Sub
    Private Sub ExcelSeriesTextBox_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ExcelSeriesTextBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub ExcelSeriesTextBox_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ExcelSeriesTextBox.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                ExcelSeriesTextBox.Text = MyFiles(i)
            Next
        End If
    End Sub
    Private Sub RadiusCenter_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RadiusCenter.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Private Sub RadiusCenter_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles RadiusCenter.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                RadiusCenter.Text = MyFiles(i)
            Next
        End If
    End Sub
    Private Sub CyclicalHeadingOption_CheckedChanged(sender As Object, e As EventArgs) Handles CyclicalHeadingOption.CheckedChanged
        LinearHeadingOption.Checked = False
    End Sub
    Private Sub CyclicalPitchOption_CheckedChanged(sender As Object, e As EventArgs) Handles CyclicalPitchOption.CheckedChanged
        LinearPitchOption.Checked = False
    End Sub
    Private Sub CyclicalRollOption_CheckedChanged_1(sender As Object, e As EventArgs) Handles CyclicalRollOption.CheckedChanged
        LinearRollOption.Checked = False
    End Sub
    Private Sub ClearPlacemarks_Click(sender As Object, e As EventArgs) Handles ClearPlacemarks.Click
        PmFolderListBox.Items.Clear()
    End Sub
    Private Sub ClearModels_Click(sender As Object, e As EventArgs) Handles ClearModels.Click
        DaeModelListBox.Items.Clear()
    End Sub
    'Private Function DistBetweenPlacemarks(lat1 As String, lon1 As String, lat2 As String, lon2 As String) As Object

    '    lat1 = lat1 * pi / 180
    '    lat2 = lat2 * pi / 180
    '    lon1 = lon1 * pi / 180
    '    lon2 = lon2 * pi / 180

    '    DistBetweenPlacemarks = Math.Acos(Math.Sin(lat1) * Math.Sin(lat2) + Math.Cos(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)) * EarthRadius
    'End Function
    'Private Function BearingBetweenPlacemarks(lat1 As String, lon1 As String, lat2 As String, lon2 As String) As Object

    '    lat1 = lat1 * pi / 180
    '    lat2 = lat2 * pi / 180
    '    lon1 = lon1 * pi / 180
    '    lon2 = lon2 * pi / 180

    '    Dim y = Math.Sin(lon2 - lon1) * Math.Cos(lat2)
    '    Dim x = Math.Cos(lat1) * Math.Sin(lat2) - Math.Sin(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)
    '    BearingBetweenPlacemarks = Math.Atan2(y, x) * 180 / pi

    'End Function
    'Private Function xfromdistbearing(dist As Object, bearing As Object) As Object
    '    xfromdistbearing = Math.Sin(bearing * pi / 180) * dist
    'End Function
    'Private Function yfromdistbearing(dist As Object, bearing As Object) As Object

    '    yfromdistbearing = Math.Cos(bearing * pi / 180) * dist
    'End Function
End Class
