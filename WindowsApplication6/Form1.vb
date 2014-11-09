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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim mystring As String = ""
        ' Load Placemarks

        NPlacemarks = ListBox1.Items.Count - 1

        Dim k As XNamespace = "http://www.opengis.net/kml/2.2"

        'XPlaceMark1 = XElement.Load(ListBox1.Items(0))
        'XPlaceMark2 = XElement.Load(ListBox1.Items(1))

        For pm = 0 To NPlacemarks
            XPlaceMark(pm) = XElement.Load(ListBox1.Items(pm))
        Next

        Dim coordinates(10) As String

        If (XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark").Count) > 0 Then
            NPlacemarks = XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark").Count - 1
            For pm = 0 To NPlacemarks
                coordinates(pm) = (XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(pm).Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)
                'MsgBox(coordinates(pm))
            Next
        Else
            For pm = 0 To NPlacemarks
                coordinates(pm) = (XPlaceMark(pm).Elements(k + "Document").Elements(k + "Placemark").Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)
            Next
        End If



        'Dim coordinates1 As String = (XPlaceMark1.Elements(k + "Document").Elements(k + "Placemark").Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)
        'Dim coordinates2 As String = (XPlaceMark2.Elements(k + "Document").Elements(k + "Placemark").Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)





        'Dim firstcomma1 As Integer = (InStr(coordinates1, ","))
        'Dim firstcomma2 As Integer = (InStr(coordinates2, ","))

        Dim firstcomma(10) As Integer

        For pm = 0 To NPlacemarks

            firstcomma(pm) = (InStr(coordinates(pm), ","))
            'MsgBox(coordinates(pm))
        Next

        'Dim longitude As Double = Microsoft.VisualBasic.Left(coordinates1, firstcomma1 - 1)
        'Dim longitudeMax As Double = Microsoft.VisualBasic.Left(coordinates2, firstcomma2 - 1)

        Dim longitudes(10) As Double

        For pm = 0 To NPlacemarks
            longitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        'coordinates1 = Microsoft.VisualBasic.Right(coordinates1, Len(coordinates1) - firstcomma1)
        'coordinates2 = Microsoft.VisualBasic.Right(coordinates2, Len(coordinates2) - firstcomma2)

        For pm = 0 To NPlacemarks
            coordinates(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        'firstcomma1 = (InStr(coordinates1, ","))
        'firstcomma2 = (InStr(coordinates2, ","))

        For pm = 0 To NPlacemarks
            firstcomma(pm) = (InStr(coordinates(pm), ","))
        Next

        'Dim latitude As Double = Microsoft.VisualBasic.Left(coordinates1, firstcomma1 - 1)
        'Dim latitudeMax As Double = Microsoft.VisualBasic.Left(coordinates2, firstcomma2 - 1)

        Dim latitudes(10) As Double

        For pm = 0 To NPlacemarks
            latitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        'Dim altitude As Double = Microsoft.VisualBasic.Right(coordinates1, Len(coordinates1) - firstcomma1)
        'Dim altitudeMax As Double = Microsoft.VisualBasic.Right(coordinates2, Len(coordinates2) - firstcomma2)

        Dim altitudes(10) As Double

        For pm = 0 To NPlacemarks
            altitudes(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        ' Load Model
        LoadModel()

        ' Write to Excel

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB As Excel.Workbook = Nothing

        Dim oCells As Excel.Range = Nothing
        oXL = New Excel.Application
        oXL.Visible = False
        oWBs = oXL.Workbooks

        oWB = oWBs.Open("C:\Google Earth Tour\Template1.xlsx")
        Dim Document As Excel.Worksheet = oWB.Worksheets(1)

        Dim Folder As Excel.Worksheet = oWB.Worksheets(2)

        Dim PlaceMark As Excel.Worksheet = oWB.Worksheets(3)
        Dim Tour As Excel.Worksheet = oWB.Worksheets(4)
        Dim FlyTo As Excel.Worksheet = oWB.Worksheets(5)

        'Tour.Range("a1").Value = TourName.Text

        Dim m2 As Date = ns1when.Value
        Dim m3 As Date = m2

        'Dim dist = DistBetweenPlacemarks(latitude, longitude, latitudeMax, longitudeMax)

        'Dim RowCount As Integer


        Dim DistanceArray(10) As Double

        DistanceArray = DistBetweenPlacemarks(latitudes, longitudes)



        'Dim bearing = BearingBetweenPlacemarks(latitude, longitude, latitudeMax, longitudeMax)

        Dim GlobalBearingArray(10) As Double

        GlobalBearingArray = GlobalBearingBetweenPlacemarks(latitudes, longitudes)

        'MsgBox(GlobalBearingArray(0) & " " & GlobalBearingArray(1) & " " & GlobalBearingArray(2))



        'Dim x = xfromdistbearing(dist, bearing)
        'Dim y = yfromdistbearing(dist, bearing)

        Dim XArray(10) As Double
        Dim YArray(10) As Double

        XArray = xarrayfromdistbearing(DistanceArray, GlobalBearingArray)
        YArray = yarrayfromdistbearing(DistanceArray, GlobalBearingArray)

        Dim LocalBearingArray(10) As Double

        LocalBearingArray = LocalBearingBetweenPlacemarks(XArray, YArray)



        'MsgBox(LocalBearingArray(0) & " " & LocalBearingArray(1) & " " & LocalBearingArray(2))

        Dim SpeedArray(10) As Double

        For speedindex = 0 To NPlacemarks
            SpeedArray(speedindex) = SpeedMin.Text + (SpeedMax.Text - SpeedMin.Text) / NPlacemarks * speedindex
        Next

        Dim TimeArray(10) As Integer

        TimeArray = TimeArrayfromDistanceArray(DistanceArray)

        'MsgBox(TimeArray(0) & " " & TimeArray(1) & " " & TimeArray(2))




        Dim VxArray(10) As Double

        VxArray = VxArrayfromLocalBearingAndSpeed(LocalBearingArray, SpeedArray)

        Dim VyArray(10) As Double

        VyArray = VyArrayfromLocalBearingAndSpeed(LocalBearingArray, SpeedArray)



        Dim AxArray(10) As Double

        AxArray = AArrayfromPositionTimeSpeed(XArray, TimeArray, VxArray)

        Dim BxArray(10) As Double

        BxArray = BArrayfromPositionTimeSpeed(XArray, TimeArray, VxArray)

        'MsgBox(BxArray(0) & " " & BxArray(1) & " " & BxArray(2))

        Dim AyArray(10) As Double

        AyArray = AArrayfromPositionTimeSpeed(YArray, TimeArray, VyArray)

        Dim ByArray(10) As Double

        ByArray = BArrayfromPositionTimeSpeed(YArray, TimeArray, VyArray)


        'For a = 0 To NPlacemarks
        '    mystring = mystring & BxArray(a) & "," & ByArray(a) & "      "
        'Next

        'MsgBox(mystring)
   

        'MsgBox(VxArray(1) & " " & VyArray(1) & " " & VxArray(2) & " " & VyArray(2))

        'MsgBox(AxArray(2) & " " & AyArray(2) & " " & BxArray(2) & " " & ByArray(2))

        'MsgBox(DistanceArray(1) & " " & XArray(1) & " " & YArray(1))



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

        Dim PMtilt As Double = PMtiltMin.Text
        Dim PMtiltMax1 As Double = PMtiltMax.Text

        Dim PMroll As Double = PMrollMin.Text
        Dim PMrollMax1 As Double = PMrollMax.Text

        Dim altitudeMode As String = ns2altitudeMode.Text
        Dim duration As Double = ns2duration.Text 'this is a percentage of the total length of tour, to the total duration.
        Dim flyToMode As String = ns2flyToMode.Text
        Dim j As Integer
        j = 0

        'For i = 0 To RowCount
        '    ProgressBar1.Value = i
        '    FlyTo.Range("a2").Offset(i, 0).Value = altitudeMode
        '    FlyTo.Range("b2").Offset(i, 0).Value = horizFov
        '    FlyTo.Range("c2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
        '    FlyTo.Range("d2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
        '    FlyTo.Range("e2").Offset(i, 0).Value = longitude + (longitudeMax - longitude) / RowCount1 * i
        '    FlyTo.Range("f2").Offset(i, 0).Value = latitude + (latitudeMax - latitude) / RowCount1 * i
        '    FlyTo.Range("g2").Offset(i, 0).Value = altitude + (altitudeMax - altitude) / RowCount1 * i
        '    FlyTo.Range("h2").Offset(i, 0).Value = heading + (headingMax - heading) / RowCount1 * i
        '    FlyTo.Range("i2").Offset(i, 0).Value = tilt + (tiltMax - tilt) / RowCount1 * i
        '    FlyTo.Range("j2").Offset(i, 0).Value = range + (rangeMax - range) / RowCount1 * i
        '    FlyTo.Range("k2").Offset(i, 0).Value = duration
        '    FlyTo.Range("l2").Offset(i, 0).Value = flyToMode

        '    PlaceMark.Range("b2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
        '    m3 = CDate(Date.FromOADate(CDbl(m3.ToOADate()) + 1 / 60 / 60 / 24 * RowCount / RowCount1))
        '    PlaceMark.Range("c2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
        '    PlaceMark.Range("e2").Offset(i, 0).Value = altitudeMode
        '    PlaceMark.Range("f2").Offset(i, 0).Value = longitude + (longitudeMax - longitude) / RowCount1 * i
        '    PlaceMark.Range("g2").Offset(i, 0).Value = latitude + (latitudeMax - latitude) / RowCount1 * i
        '    PlaceMark.Range("h2").Offset(i, 0).Value = altitude + (altitudeMax - altitude) / RowCount1 * i
        '    PlaceMark.Range("i2").Offset(i, 0).Value = PMheading + (PMheadingMax1 - PMheading) / RowCount1 * i

        '    PlaceMark.Range("j2").Offset(i, 0).Value = PMtilt + (PMtiltMax1 - PMtilt) / RowCount1 * i
        '    PlaceMark.Range("k2").Offset(i, 0).Value = PMroll + (PMrollMax1 - PMroll) / RowCount1 * i


        '    PlaceMark.Range("o2").Offset(i, 0).Value = DaeName(j)
        '    If j = DaeNameSteps Then j = 0 Else j = j + 1

        'Next

        Dim xPosition As Double = 0
        Dim yPosition As Double = 0
        Dim DistanceBetweenXY As Double = 0
        Dim BearingBetweenXY As Double = 0

        Dim OutputLatDeg As Double = latitudes(0) * 180 / pi
        Dim OutputLongDeg As Double = longitudes(0) * 180 / pi

        'Dim RowCount1

        Dim ModelX As Double
        Dim ModelY As Double
        Dim ModelBearing As Double

        Dim OutputLatDegPrevious As Double
        Dim OutputLongDegPrevious As Double

        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = TimeArray(NPlacemarks)

        ProgressBar1.Visible = True
        ProgressBar1.Value = ProgressBar1.Minimum

        Dim f As Integer = 0
        Dim i As Double = 0


        'For h = 1 To NPlacemarks
        '    For i = TimeArray(h - 1) To TimeArray(h) - 1

        For h = 1 To NPlacemarks

            While i < TimeArray(h)

                xPosition = 1 / 6 * AxArray(h) * (i - TimeArray(h - 1)) ^ 3 + 1 / 2 * BxArray(h) * (i - TimeArray(h - 1)) ^ 2 + VxArray(h - 1) * (i - TimeArray(h - 1)) + XArray(h - 1) '+XArray(0) 

                'MsgBox(AxArray(h) & " " & TimeArray(h - 1) & " " & BxArray(h) & " " & VxArray(h) & " " & XArray(h - 1))

                yPosition = 1 / 6 * AyArray(h) * (i - TimeArray(h - 1)) ^ 3 + 1 / 2 * ByArray(h) * (i - TimeArray(h - 1)) ^ 2 + VyArray(h - 1) * (i - TimeArray(h - 1)) + YArray(h - 1) '+ YArray(0)

               
                'MsgBox(AyArray(h) & " " & TimeArray(h - 1) & " " & ByArray(h) & " " & VyArray(h) & " " & YArray(h - 1))

                DistanceBetweenXY = (xPosition ^ 2 + yPosition ^ 2) ^ 0.5

                BearingBetweenXY = Math.Atan2(yPosition, xPosition) - 90 * pi / 180

                'If i = TimeArray(h - 1) Then

                '    mystring = i & "," & DistanceBetweenXY & "," & BearingBetweenXY & "     "

                '    MsgBox(mystring)
                'End If

                OutputLatDegPrevious = OutputLatDeg
                'OutputLatDeg = (Math.Asin(Math.Sin(latitudes(h - 1)) * Math.Cos(DistanceBetweenXY / EarthRadius) + Math.Cos(latitudes(h - 1)) * Math.Sin(DistanceBetweenXY / EarthRadius) * Math.Cos(BearingBetweenXY))) * 180 / pi
                OutputLatDeg = (Math.Asin(Math.Sin(latitudes(0)) * Math.Cos(DistanceBetweenXY / 1000 / 6378.1) + Math.Cos(latitudes(0)) * Math.Sin(DistanceBetweenXY / 1000 / 6378.1) * Math.Cos(BearingBetweenXY))) * 180 / pi
                OutputLongDegPrevious = OutputLongDeg
                'OutputLongDeg = (longitudes(h - 1) + Math.Atan2(Math.Cos(DistanceBetweenXY / EarthRadius) - Math.Sin(latitudes(h - 1)) * Math.Sin(OutputLatDeg * pi / 180), Math.Sin(BearingBetweenXY) * Math.Sin(DistanceBetweenXY / EarthRadius) * Math.Cos(latitudes(h - 1)))) * 180 / pi - 90
                OutputLongDeg = (longitudes(0) + Math.Atan2(Math.Cos(DistanceBetweenXY / EarthRadius) - Math.Sin(latitudes(0)) * Math.Sin(OutputLatDeg * pi / 180), Math.Sin(BearingBetweenXY) * Math.Sin(DistanceBetweenXY / EarthRadius) * Math.Cos(latitudes(0)))) * 180 / pi - 90




                'MsgBox(t & " " & OutputLatDeg & " " & OutputLongDeg)

                ProgressBar1.Value = i

                'Set the view

                FlyTo.Range("a2").Offset(i, 0).Value = altitudeMode
                FlyTo.Range("b2").Offset(i, 0).Value = horizFov
                FlyTo.Range("c2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                FlyTo.Range("d2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                FlyTo.Range("e2").Offset(i, 0).Value = OutputLongDeg 'longitude + (longitudeMax - longitude) / RowCount1 * i
                FlyTo.Range("f2").Offset(i, 0).Value = OutputLatDeg 'latitude + (latitudeMax - latitude) / RowCount1 * i
                FlyTo.Range("g2").Offset(i, 0).Value = altitudes(0) '+ (altitudeMax - altitude) / RowCount1 * i
                FlyTo.Range("h2").Offset(i, 0).Value = heading + (headingMax - heading) / TimeArray(NPlacemarks) * f
                FlyTo.Range("i2").Offset(i, 0).Value = tilt + (tiltMax - tilt) / TimeArray(NPlacemarks) * f
                FlyTo.Range("j2").Offset(i, 0).Value = range + (rangeMax - range) / TimeArray(NPlacemarks) * f
                FlyTo.Range("k2").Offset(i, 0).Value = duration
                FlyTo.Range("l2").Offset(i, 0).Value = flyToMode

                'Set the model

                PlaceMark.Range("b2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                m3 = CDate(Date.FromOADate(CDbl(m3.ToOADate()) + 1 / 60 / 60 / 24))
                'm3 = CDate(Date.FromOADate(CDbl(m2.ToOADate()) + 1 / 60 / 60 / 24))
                PlaceMark.Range("c2").Offset(i, 0).Value = Year(m3) & "-" & Format(Month(m3), "00") & "-" & Format(Day(m3), "00") & "T" & Format(Hour(m3), "00") & ":" & Format(Minute(m3), "00") & ":" & Format(Second(m3), "00") & "Z"
                PlaceMark.Range("e2").Offset(i, 0).Value = altitudeMode
                PlaceMark.Range("f2").Offset(i, 0).Value = OutputLongDeg 'longitude + (longitudeMax - longitude) / RowCount1 * i
                PlaceMark.Range("g2").Offset(i, 0).Value = OutputLatDeg 'latitude + (latitudeMax - latitude) / RowCount1 * i
                PlaceMark.Range("h2").Offset(i, 0).Value = altitudes(0) '+ (altitudeMax - altitude) / RowCount1 * i


                ModelY = Math.Sin((OutputLongDeg - OutputLongDegPrevious) * pi / 180) * Math.Cos(OutputLatDeg * pi / 180)
                ModelX = Math.Cos(OutputLatDegPrevious * pi / 180) * Math.Sin(OutputLatDeg * pi / 180) - Math.Sin(OutputLatDegPrevious * pi / 180) * Math.Cos(OutputLatDeg * pi / 180) * Math.Cos((OutputLongDeg - OutputLongDegPrevious) * pi / 180)
                ModelBearing = Math.Atan2(ModelY, ModelX) * 180 / pi - 90
                'MsgBox(ModelBearing)

                PlaceMark.Range("i2").Offset(i, 0).Value = ModelBearing 'PMheading '+ (PMheadingMax1 - PMheading) / RowCount1 * i

                PlaceMark.Range("j2").Offset(i, 0).Value = PMtilt + (PMtiltMax1 - PMtilt) / TimeArray(NPlacemarks) * f
                PlaceMark.Range("k2").Offset(i, 0).Value = PMroll + (PMrollMax1 - PMroll) / TimeArray(NPlacemarks) * f


                PlaceMark.Range("o2").Offset(i, 0).Value = DaeName(j)
                If j = DaeNameSteps Then j = 0 Else j = j + 1
                f = f + 1
                i = i + 1

            End While
        Next

        'Next

        'Dim MyPath As String = "C:\Resource Documents\Resources\Simulation\Google Earth Animation\"
        Dim MyPath As String = "C:\Google Earth Tour\"
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

        'ProgressBar1.Visible = False
        ProgressBar1.Value = ProgressBar1.Minimum

    End Sub

    Private MouseIsDown As Boolean = False

    Private Sub ListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub ListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                ListBox1.Items.Add(MyFiles(i))
            Next
        End If
    End Sub
    Private Sub ListBox2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox2.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub ListBox2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox2.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer

            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                ListBox2.Items.Add(MyFiles(i))
            Next
        End If
    End Sub

    Private Sub LoadModel()

        DaeNameSteps = -1

        For k = 0 To ListBox2.Items.Count - 1
            DaeName(k) = ListBox2.Items(k)
            DaeNameSteps = DaeNameSteps + 1
        Next
    End Sub

    Private Sub ClearPlacemarks_Click(sender As Object, e As EventArgs) Handles ClearPlacemarks.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub ClearModels_Click(sender As Object, e As EventArgs) Handles ClearModels.Click
        ListBox2.Items.Clear()
    End Sub

    Private Function DistBetweenPlacemarks(lat1 As String, lon1 As String, lat2 As String, lon2 As String) As Object

        lat1 = lat1 * pi / 180
        lat2 = lat2 * pi / 180
        lon1 = lon1 * pi / 180
        lon2 = lon2 * pi / 180

        DistBetweenPlacemarks = Math.Acos(Math.Sin(lat1) * Math.Sin(lat2) + Math.Cos(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)) * EarthRadius
    End Function

    Private Function BearingBetweenPlacemarks(lat1 As String, lon1 As String, lat2 As String, lon2 As String) As Object

        lat1 = lat1 * pi / 180
        lat2 = lat2 * pi / 180
        lon1 = lon1 * pi / 180
        lon2 = lon2 * pi / 180

        Dim y = Math.Sin(lon2 - lon1) * Math.Cos(lat2)
        Dim x = Math.Cos(lat1) * Math.Sin(lat2) - Math.Sin(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)
        BearingBetweenPlacemarks = Math.Atan2(y, x) * 180 / pi

    End Function

    Private Function xfromdistbearing(dist As Object, bearing As Object) As Object
        xfromdistbearing = Math.Sin(bearing * pi / 180) * dist
        End Function
    Private Function yfromdistbearing(dist As Object, bearing As Object) As Object

        yfromdistbearing = Math.Cos(bearing * pi / 180) * dist
    End Function

    Private Function DistBetweenPlacemarks(latitudes As Double(), longitudes As Double()) As Double()
        Dim lat1 As Double
        Dim lat2 As Double
        Dim lon1 As Double
        Dim lon2 As Double
        Dim ArrayHolder(10) As Double

        For k = 0 To NPlacemarks

            lat1 = latitudes(0)
            lat2 = latitudes(k)
            lon1 = longitudes(0)
            lon2 = longitudes(k)

            ArrayHolder(k) = Math.Acos(Math.Sin(lat1) * Math.Sin(lat2) + Math.Cos(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)) * EarthRadius
        Next
        Return ArrayHolder
    End Function

    Private Function GlobalBearingBetweenPlacemarks(latitudes As Double(), longitudes As Double()) As Double()
        Dim lat1 As Double
        Dim lat2 As Double
        Dim lon1 As Double
        Dim lon2 As Double
        Dim ArrayHolder(10) As Double

        For k = 1 To NPlacemarks

            lat1 = latitudes(0)
            lat2 = latitudes(k)
            lon1 = longitudes(0)
            lon2 = longitudes(k)

            Dim y = Math.Sin(lon2 - lon1) * Math.Cos(lat2)
            Dim x = Math.Cos(lat1) * Math.Sin(lat2) - Math.Sin(lat1) * Math.Cos(lat2) * Math.Cos(lon2 - lon1)
            ArrayHolder(k) = Math.Atan2(y, x)

        Next
        Return ArrayHolder
    End Function

    Private Function LocalBearingBetweenPlacemarks(XArray As Double(), YArray As Double()) As Double()
        Dim x1 As Double
        Dim x2 As Double
        Dim y1 As Double
        Dim y2 As Double
        Dim ArrayHolder(10) As Double

        For k = 1 To NPlacemarks

            x1 = XArray(k - 1)
            x2 = XArray(k)
            y1 = YArray(k - 1)
            y2 = YArray(k)

            Dim y = (y2 - y1)
            Dim x = (x2 - x1)
            ArrayHolder(k) = Math.Atan2(y, x)

        Next
        Return ArrayHolder
    End Function


    Private Function xarrayfromdistbearing(DistanceArray As Double(), BearingArray As Double()) As Double()
        Dim ArrayHolder(10) As Double
        For k = 0 To NPlacemarks
            ArrayHolder(k) = Math.Sin(BearingArray(k)) * DistanceArray(k)
        Next
        Return ArrayHolder
    End Function

    Private Function yarrayfromdistbearing(DistanceArray As Double(), BearingArray As Double()) As Double()
        Dim ArrayHolder(10) As Double
        For k = 0 To NPlacemarks
            ArrayHolder(k) = Math.Cos(BearingArray(k)) * DistanceArray(k)
        Next
        Return ArrayHolder
    End Function

    Private Function TimeArrayfromDistanceArray(DistanceArray As Double()) As Integer()
        Dim ArrayHolder(10) As Integer
        ArrayHolder(0) = 100 * DistanceArray(0) / 120

        For k = 1 To NPlacemarks
            ArrayHolder(k) = 100 * DistanceArray(k) / 120 + ArrayHolder(k - 1)
        Next

        Return ArrayHolder
    End Function

    Private Function VxArrayfromLocalBearingAndSpeed(LocalBearingArray As Double(), SpeedArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For k = 0 To NPlacemarks
            ArrayHolder(k) = SpeedArray(k) * Math.Cos(LocalBearingArray(k))
            'MsgBox(ArrayHolder(k))
        Next
        Return ArrayHolder
    End Function

    Private Function VyArrayfromLocalBearingAndSpeed(LocalBearingArray As Double(), SpeedArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For k = 0 To NPlacemarks
            ArrayHolder(k) = SpeedArray(k) * Math.Sin(LocalBearingArray(k))
        Next
        Return ArrayHolder
    End Function

    Private Function AArrayfromPositionTimeSpeed(XYArray As Double(), TimeArray As Integer(), VxyArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For k = 1 To NPlacemarks
            ArrayHolder(k) = 6 * ((VxyArray(k) + VxyArray(k - 1)) * (TimeArray(k) - TimeArray(k - 1)) - 2 * (XYArray(k) - XYArray(k - 1))) / (TimeArray(k) - TimeArray(k - 1)) ^ 3
        Next
        Return ArrayHolder
    End Function

    Private Function BArrayfromPositionTimeSpeed(XYArray As Double(), TimeArray As Integer(), VxyArray As Double()) As Double()
        Dim ArrayHolder(10) As Double

        For k = 1 To NPlacemarks
            ArrayHolder(k) = -2 * ((VxyArray(k) + 2 * VxyArray(k - 1)) * (TimeArray(k) - TimeArray(k - 1)) - 3 * (XYArray(k) - XYArray(k - 1))) / (TimeArray(k) - TimeArray(k - 1)) ^ 2

        Next
        Return ArrayHolder
    End Function


End Class
