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
    Dim XPlacemark_Data As XElement = XElement.Load("C:\Google Earth Tour\PlacemarkDataTemplate.xml")
    Dim XAnimateModel As XElement = XElement.Load("C:\Google Earth Tour\AnimateModelTemplate.xml")
    Dim XTrack As XElement = XElement.Load("C:\Google Earth Tour\TrackTemplate.xml")
    Dim DaeName(8) As String
    Dim DaeNameSteps As Integer
    Dim pi = 3.14159265358979
    Dim EarthRadius = 6378.1 * 1000
    Dim NPlacemarks As Integer
    Dim k As XNamespace = "http://www.opengis.net/kml/2.2"
    Dim kk As XNamespace = "http://www.google.com/kml/ext/2.2"

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
        oXL = New Excel.Application
        oXL.Visible = True

        Dim oWBs As Excel.Workbooks = Nothing

        oWBs = oXL.Workbooks

        Dim oWB2 As Excel.Workbook = Nothing

        Dim Sheet1 As Excel.Worksheet

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

        Dim BeginTime As String
        Dim EndTime As String
        Dim OutputString As String

        Dim CoordinateString As String
        Dim OrientationString As String
        Dim TiltString As String
        Dim RangeString As String

        Dim TrimData As Double
        Dim HeelData As Double
        Dim SpeedData As Double


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

                BeginTime = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                
                CoordinateString = OutputLongDeg & " " & OutputLatDeg & " " & altitudes(0)

                If LinearHeadingOption.Checked = True Then
                    OrientationString = heading + (headingMax - heading) / TimeArray(NPlacemarks) * i
                Else
                    OrientationString = i Mod 360
                End If

                TiltString = tilt + (tiltMax - tilt) / TimeArray(NPlacemarks) * i
                RangeString = range + (rangeMax - range) / TimeArray(NPlacemarks) * i

                'Set the model

                m3 = CDate(Date.FromOADate(CDbl(m3.ToOADate()) + TimeIncrementText / 60 / 60 / 24))

                EndTime = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"

                OutputString = OutputLongDeg & "," & OutputLatDeg & "," & altitudes(0)

                ModelY = Math.Sin((OutputLongDeg - OutputLongDegPrevious) * pi / 180) * Math.Cos(OutputLatDeg * pi / 180)
                ModelX = Math.Cos(OutputLatDegPrevious * pi / 180) * Math.Sin(OutputLatDeg * pi / 180) - Math.Sin(OutputLatDegPrevious * pi / 180) * Math.Cos(OutputLatDeg * pi / 180) * Math.Cos((OutputLongDeg - OutputLongDegPrevious) * pi / 180)
                ModelBearing = Math.Atan2(ModelY, ModelX) * 180 / pi - 90

                If LinearRollOption.Checked = True Then
                    TrimData = PMtilt + (PMtiltMax1 - PMtilt) / TimeArray(NPlacemarks) * i
                Else
                    TrimData = RollMagnitude.Text * Math.Sin(2 * pi / RollPeriod.Text * i + RollPhase.Text * pi / 180)
                End If

                TrimString = "Trim: " & Math.Round(TrimData, 1) & "°; "

                If LinearPitchOption.Checked = True Then
                    HeelData = PMroll + (PMrollMax1 - PMroll) / TimeArray(NPlacemarks) * i

                Else
                    HeelData = PitchMagnitude.Text * Math.Sin(2 * pi / PitchPeriod.Text * i + PitchPhase.Text * pi / 180)
                End If

                HeelString = "Heel: " & Math.Round(HeelData, 1) & "°; "

                If j = DaeNameSteps Then j = 0 Else j = j + 1

                'MsgBox((ModelX ^ 2 + ModelY ^ 2) ^ 0.5)

                HeadingString = "Heading: " & Math.Round(ModelBearing + 90, 1) & "°; "

                SpeedData = Math.Acos(Math.Sin(OutputLatDegPrevious * pi / 180) * Math.Sin(OutputLatDeg * pi / 180) + Math.Cos(OutputLatDegPrevious * pi / 180) * Math.Cos(OutputLatDeg * pi / 180) * Math.Cos(OutputLongDeg * pi / 180 - OutputLongDegPrevious * pi / 180)) * EarthRadius

                SpeedString = "Speed: " & Math.Round(SpeedData / TimeIncrementText, 1) & "m/s; "
                DraftString = "Draft: " & Math.Round(altitudes(0), 1) & "m ; "

                AddToPlacemarkData(HeadingString & SpeedString & DraftString & TrimString & HeelString, BeginTime, EndTime, OutputString, index)
                AddToAnimateModel(altitudeMode, horizFov, BeginTime, OutputLongDeg, OutputLatDeg, altitudes(0), OrientationString, TiltString, RangeString, duration, flyToMode, EndTime, ModelBearing, DaeName(j), TrimData, HeelData, index)
                AddToTrack(BeginTime, CoordinateString, index)

                i = i + TimeIncrementText
                index = index + 1

            End While
        Next

        XPlacemark_Data.Save("C:\Google Earth Tour\PlacemarkData" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")
        XTrack.Save("C:\Google Earth Tour\Track" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")
        XAnimateModel.Save("C:\Google Earth Tour\Model" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")

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

    Private Sub AddToPlacemarkData(HeadingString As String, BeginTime As String, EndTime As String, OutputString As String, index As Integer)
        Dim xAdd As XElement
        xAdd = <ns1:Placemark id="pm267" xmlns:ns1="http://www.opengis.net/kml/2.2">
                   <ns1:name><%= HeadingString %></ns1:name>
                   <ns1:Snippet maxLines="0">empty</ns1:Snippet>
                   <ns1:description>hello</ns1:description>
                   <ns1:TimeSpan>
                       <ns1:begin><%= BeginTime %></ns1:begin>
                       <ns1:end><%= EndTime %></ns1:end>
                   </ns1:TimeSpan>
                   <ns1:styleUrl>#Style_5</ns1:styleUrl>
                   <ns1:Point>
                       <ns1:coordinates><%= OutputString %></ns1:coordinates>
                   </ns1:Point>
               </ns1:Placemark>

        If index = 0 Then
            XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index).ReplaceWith(xAdd)
        Else
            XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index - 1).AddAfterSelf(xAdd)
        End If

    End Sub

    Private Sub AddToAnimateModel(altitudeMode As String, horizFov As String, BeginTime As String, OutputLongDeg As Double, OutputLatDeg As Double, altitudes As Double, OrientationString As String, TiltString As String, RangeString As String, duration As Double, flyToMode As String, EndTime As String, ModelBearing As Double, DaeName As String, TrimData As Double, HeelData As Double, index As Integer)
        Dim xPlacemarkTable As XElement
        Dim xFlytoTable As XElement

        xFlytoTable = <ns2:FlyTo xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                          <ns1:LookAt>
                              <ns2:altitudeMode><%= altitudeMode %></ns2:altitudeMode>
                              <ns2:horizFov><%= horizFov %></ns2:horizFov>
                              <ns2:TimeSpan>
                                  <ns1:begin><%= BeginTime %></ns1:begin>
                                  <ns1:end><%= EndTime %></ns1:end>
                              </ns2:TimeSpan>
                              <ns1:longitude><%= OutputLongDeg %></ns1:longitude>
                              <ns1:latitude><%= OutputLatDeg %></ns1:latitude>
                              <ns1:altitude><%= altitudes %></ns1:altitude>
                              <ns1:heading><%= OrientationString %></ns1:heading>
                              <ns1:tilt><%= TiltString %></ns1:tilt>
                              <ns1:range><%= RangeString %></ns1:range>
                          </ns1:LookAt>
                          <ns2:duration><%= duration %></ns2:duration>
                          <ns2:flyToMode><%= flyToMode %></ns2:flyToMode>
                      </ns2:FlyTo>

        xPlacemarkTable = <ns1:Placemark xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                              <ns1:name>1</ns1:name>
                              <ns1:TimeSpan>
                                  <ns1:begin><%= BeginTime %></ns1:begin>
                                  <ns1:end><%= EndTime %></ns1:end>
                              </ns1:TimeSpan>
                              <ns1:MultiGeometry>
                                  <ns1:Model id="model_1">
                                      <ns1:altitudeMode><%= altitudeMode %></ns1:altitudeMode>
                                      <ns1:Location>
                                          <ns1:longitude><%= OutputLongDeg %></ns1:longitude>
                                          <ns1:latitude><%= OutputLatDeg %></ns1:latitude>
                                          <ns1:altitude><%= altitudes %></ns1:altitude>
                                      </ns1:Location>
                                      <ns1:Orientation>
                                          <ns1:heading><%= ModelBearing %></ns1:heading>
                                          <ns1:tilt><%= TrimData %></ns1:tilt>
                                          <ns1:roll><%= HeelData %></ns1:roll>
                                      </ns1:Orientation>
                                      <ns1:Scale>
                                          <ns1:x>1</ns1:x>
                                          <ns1:y>1</ns1:y>
                                          <ns1:z>1</ns1:z>
                                      </ns1:Scale>
                                      <ns1:Link>
                                          <ns1:href><%= DaeName %></ns1:href>
                                      </ns1:Link>
                                  </ns1:Model>
                              </ns1:MultiGeometry>
                          </ns1:Placemark>

        If index = 0 Then
            XAnimateModel.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index).ReplaceWith(xPlacemarkTable)
            XAnimateModel.Elements(k + "Document").Elements(kk + "Tour").Elements(kk + "Playlist").Elements(kk + "FlyTo")(index).ReplaceWith(xFlytoTable)

        Else
            XAnimateModel.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index - 1).AddAfterSelf(xPlacemarkTable)
            XAnimateModel.Elements(k + "Document").Elements(kk + "Tour").Elements(kk + "Playlist").Elements(kk + "FlyTo")(index - 1).AddAfterSelf(xFlytoTable)

        End If


    End Sub



    Private Sub AddToTrack(BeginTime As String, CoordinateString As String, index As Integer)
        Dim xAddTime As XElement
        Dim xAddCoord As XElement
        xAddTime = <ns1:when xmlns:ns1="http://www.opengis.net/kml/2.2"><%= BeginTime %></ns1:when>
        xAddCoord = <ns2:coord xmlns:ns2="http://www.google.com/kml/ext/2.2"><%= CoordinateString %></ns2:coord>

        If index = 0 Then
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(k + "when")(index).ReplaceWith(xAddTime)
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(kk + "coord")(index).ReplaceWith(xAddCoord)
        Else
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(k + "when")(index - 1).AddAfterSelf(xAddTime)
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(kk + "coord")(index - 1).AddAfterSelf(xAddCoord)
        End If

    End Sub

End Class
