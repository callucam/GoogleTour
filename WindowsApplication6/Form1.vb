#Region "Imports directives"

Imports System.Reflection
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Net.Configuration

#End Region


Public Class Form1

    Dim XPlaceMark(5000) As XElement
    Dim XLookAtPlaceMark(5000) As XElement
    Dim XPlacemark_Data As XElement
    Dim XAnimateModel As XElement
    Dim XTrack As XElement
    Dim DaeName(24) As String
    Dim DaeNameSteps As Integer
    Dim pi = 3.14159265358979
    Dim EarthRadius = 6378.1 * 1000
    Dim NPlacemarks As Integer
    Dim k As XNamespace = "http://www.opengis.net/kml/2.2"
    Dim kk As XNamespace = "http://www.google.com/kml/ext/2.2"


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim mystring As String = ""
        Dim ESS = 3500
        XPlacemark_Data = XElement.Load("C:\Google Earth Tour\PlacemarkDataTemplate.xml")
        XAnimateModel = XElement.Load("C:\Google Earth Tour\AnimateModelTemplate.xml")
        XTrack = XElement.Load("C:\Google Earth Tour\TrackTemplate.xml")

        ' Set view properties

        Dim horizFov As String = ns2horizFov.Text
        Dim heading As Double = ns1heading.Text
        Dim headingMax As Double = ns1headingMax.Text
        Dim tilt As Double = ns1tilt.Text
        Dim tiltMax As Double = ns1tiltMax.Text
        Dim range As Double = ns1range.Text
        Dim rangeMax As Double = ns1rangeMax.Text

        ' Set model properties

        Dim altitudeMode As String = ns2altitudeMode.Text
        Dim duration As Double = ns2duration.Text 'this is a percentage of the total length of tour, to the total duration.
        Dim flyToMode As String = ns2flyToMode.Text

        ' Load Placemarks

        Dim LonLatAlt
        Dim longitudes(5000) As Double
        Dim latitudes(5000) As Double
        Dim altitudes(5000) As Double

        'LonLatAlt = LoadPlacemarks(PmReferenceTextBox.Text)

        'longitudes = LonLatAlt(0)
        'latitudes = LonLatAlt(1)
        'altitudes = LonLatAlt(2)

        LoadModel()

        ' Write to Excel

        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB2 As Excel.Workbook = Nothing
        Dim Sheet1 As Excel.Worksheet
        Dim m3 As Date

#Disable Warning BC42024 ' Unused local variable

        Dim m7 As Date

#Enable Warning BC42024 ' Unused local variable


        Dim DistanceArray(5000) As Double
        Dim CummulativeDistanceArray(5000) As Double

        Dim GlobalBearingArray(5000) As Double
        Dim XArray(5000) As Double
        Dim YArray(5000) As Double

        Dim HeelArray(5000) As Double
        Dim TrimArray(5000) As Double
        Dim YawArray(5000) As Double
        Dim DraftArray(5000) As Double

        Dim LocalBearingArray(5000) As Double
        Dim OrientationArray(5000) As Double

        Dim SpeedArray(5000) As Double
        Dim TimeArray(5000) As Double
        Dim VxArray(5000) As Double
        Dim VyArray(5000) As Double
        Dim AxArray(5000) As Double
        Dim BxArray(5000) As Double
        Dim AyArray(5000) As Double
        Dim ByArray(5000) As Double
        Dim j As Integer = 0
        Dim DistanceBetweenXY As Double = 0
        Dim BearingBetweenXY As Double = 0

        Dim OutputAltitude As Double

        Dim String1(5000) As String
        Dim String2(5000) As String
        Dim String3(5000) As String
        Dim String4(5000) As String
        Dim String5(5000) As String
        Dim String6(5000) As String
        Dim String7(5000) As String


        Dim ResolutionArray(5000) As Integer
        'Dim IndexArray(5000) As Integer


        Dim LatArray(5000) As String
        Dim LongArray(5000) As String

        Dim m4 As Date
        Dim m5 As Date

        Dim ReadoutSubTitleString As String
        Dim ReadoutTitleString As String


        oXL = New Excel.Application
        oXL.Visible = False
        oWBs = oXL.Workbooks
        oWB2 = oWBs.Open(ExcelSeriesTextBox.Text)
        Sheet1 = oWB2.Worksheets(1)

        ReadoutSubTitleString = Sheet1.Range("b6").Offset(0, 0).Value
        ReadoutTitleString = Sheet1.Range("b7").Offset(0, 0).Value

        NPlacemarks = Sheet1.Range("b2").Offset(0, 0).Value - 1
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = NPlacemarks
        m3 = DateAdd(DateInterval.Second, Sheet1.Range("a11").Offset(0, 0).Value, ns1when.Value)
        m3 = DateAdd(DateInterval.Hour, NumericUpDown1.Value, m3)
        m3 = DateAdd(DateInterval.Minute, NumericUpDown2.Value, m3)
        For n = 0 To NPlacemarks
            m5 = DateAdd(DateInterval.Second, Sheet1.Range("a11").Offset(n, 0).Value, ns1when.Value)
            m5 = DateAdd(DateInterval.Hour, NumericUpDown1.Value, m5)
            m5 = DateAdd(DateInterval.Minute, NumericUpDown2.Value, m5)
            TimeArray(n) = Sheet1.Range("a11").Offset(n, 0).Value 'DateDiff(DateInterval.Second, m3, m5)
            'MsgBox(m5 & " " & TimeArray(n) & " " & Sheet1.Range("a11").Offset(n, 0).Value)
            XArray(n) = Sheet1.Range("b11").Offset(n, 0).Value ' xarrayfromdistbearing(DistanceArray, GlobalBearingArray)
            YArray(n) = Sheet1.Range("c11").Offset(n, 0).Value
            HeelArray(n) = Sheet1.Range("d11").Offset(n, 0).Value
            TrimArray(n) = Sheet1.Range("e11").Offset(n, 0).Value
            YawArray(n) = Sheet1.Range("f11").Offset(n, 0).Value
            DraftArray(n) = Sheet1.Range("g11").Offset(n, 0).Value
            GlobalBearingArray(n) = Sheet1.Range("h11").Offset(n, 0).Value
            LocalBearingArray(n) = Sheet1.Range("i11").Offset(n, 0).Value
            DistanceArray(n) = Sheet1.Range("j11").Offset(n, 0).Value
            SpeedArray(n) = Sheet1.Range("k11").Offset(n, 0).Value

            String1(n) = Sheet1.Range("l11").Offset(n, 0).Value
            String2(n) = Sheet1.Range("m11").Offset(n, 0).Value
            String3(n) = Sheet1.Range("n11").Offset(n, 0).Value
            String4(n) = Sheet1.Range("y11").Offset(n, 0).Value
            String5(n) = Sheet1.Range("z11").Offset(n, 0).Value
            String6(n) = Sheet1.Range("aa11").Offset(n, 0).Value
            String7(n) = Sheet1.Range("ab11").Offset(n, 0).Value

            ResolutionArray(n) = Sheet1.Range("ac11").Offset(n, 0).Value

            'IndexArray(n) = Sheet1.Range("ae11").Offset(n, 0).Value


            CummulativeDistanceArray(n) = Sheet1.Range("o11").Offset(n, 0).Value

            LatArray(n) = Sheet1.Range("r11").Offset(n, 0).Value
            LongArray(n) = Sheet1.Range("s11").Offset(n, 0).Value
            ProgressBar1.Visible = True
            ProgressBar1.Value = n

        Next

        oWB2.Close(False)


        OrientationArray = LocalBearingArray

        m4 = DateAdd(DateInterval.Second, TimeArray(NPlacemarks), m3)

        'MsgBox(TimeArray(NPlacemarks) & " " & NPlacemarks)

        Dim VeryEndTime As String = kmlDate(m4)

        'MsgBox(VeryEndTime)

        Dim ModelBearing As Double
        Dim ModelBearingPrevious(20) As Double

        Dim i As Double = 0
        Dim index As Integer = 0

        Dim BeginTime As String
        Dim EndTime As String
        Dim OutputString As String

        Dim CoordinateString As String
        Dim OrientationString As String
        Dim TiltString As String
        Dim RangeString As String
        Dim mb As Integer = 0

        Dim ScaleData As Double

        'Dim indexForPlacemarkData As Integer = 0
        Dim ReadoutString As String
        Dim LongPos As Double
        Dim LatPos As Double
        Dim Resolution As Integer = 8

        For h = 1 To NPlacemarks

            ProgressBar2.Minimum = 0
            ProgressBar2.Maximum = NPlacemarks
            ProgressBar2.Value = h

            Resolution = ResolutionArray(h - 1)
            'MsgBox(Resolution)

            For i = 0 To Resolution - 1
                'For i = 1 To 2

                'MsgBox(h & " " & TimeArray(h) & " " & i)

                OutputAltitude = altitudes(0)

                BeginTime = kmlDate(m3)

                LongPos = LongArray(h - 1) + (LongArray(h) - LongArray(h - 1)) / (Resolution) * (i)
                LatPos = LatArray(h - 1) + (LatArray(h) - LatArray(h - 1)) / (Resolution) * (i)


                'CoordinateString = LongArray(h) & " " & LatArray(h) & " " & OutputAltitude
                CoordinateString = LongPos & " " & LatPos & " " & OutputAltitude

                'MsgBox(CoordinateString)

                If LinearHeadingOption.Checked = True Then
                    OrientationString = heading + (headingMax - heading) / ((NPlacemarks - 1) * Resolution) * (((h - 1) * Resolution) + i)
                Else
                    OrientationString = h Mod 360
                    'MsgBox("here")
                End If

                TiltString = tilt + (tiltMax - tilt) / ((NPlacemarks - 1) * Resolution) * (((h - 1) * Resolution) + i)
                RangeString = range + (rangeMax - range) / ((NPlacemarks - 1) * Resolution) * (((h - 1) * Resolution) + i)

                'Set the model

                'm3 = m3.AddMilliseconds(1 * 1000 / Resolution)
                m3 = m3.AddMilliseconds(1000 * ((TimeArray(h) - TimeArray(h - 1)) / Resolution))

                'MsgBox(TimeArray(h) & " " & TimeArray(h - 1) & " " & Resolution)

                EndTime = kmlDate(m3)

                OutputString = LongPos & "," & LatPos  'altitudes(0)

                ModelBearing = GlobalBearingArray(h - 1) + (GlobalBearingArray(h) - GlobalBearingArray(h - 1)) / Resolution * i - 90

                'PowerString = "Power"
                'BatteryString = "Battery" & i
                ScaleData = 1

                'PowerData = If(SpeedData = 0, 0, Math.Round(164.85 * Math.Exp(0.1616 * SpeedData), 0))
                'ESS = ESS - PowerData / 60 / 60

                'PassengerString = "Passengers: " & 146
                'WeatherString = "Weather: Winds SE 10 knots"
                'PowerString = "Propulsion Power: " & PowerData & " kW" ' correlated to speed in m/s
                'BatteryString = "Vessel SOC: " & Math.Round(ESS, 0) & " kWh"

                'If j = DaeNameSteps Then j = 0 Else j = j + 1

                ReadoutString = ""

                'MsgBox(OutputLongDeg & " " & OutputLatDeg & " " & LongArray(h) & " " & LatArray(h))

                If ReadoutCheckBox.Checked = True Then
                    CreateImageReadout(BeginTime, VeryEndTime, LongPos, LatPos, OrientationString, TiltString, RangeString, index, String5(h), String4(h), String2(h), String3(h), String6(h), String7(h), ReadoutTitleString, ReadoutSubTitleString)
                End If

                AddToPlacemarkData(ReadoutString, BeginTime, EndTime, OutputString, LongPos, LatPos, LatPos, OrientationString, TiltString, RangeString, index)
                AddToAnimateModel(altitudeMode, horizFov, BeginTime, LongPos, LatPos, DraftArray(h), OrientationString, TiltString, RangeString, duration, flyToMode, EndTime, ModelBearing, DaeName(j), HeelArray(h), TrimArray(h), YawArray(h), index, ScaleData)
                AddToTrack(BeginTime, CoordinateString, LongPos, LatPos, altitudes(0), OrientationString, TiltString, RangeString, index)

                'indexForPlacemarkData = indexForPlacemarkData + 1
                'i = i + 1
                'i = (((h - 1) * 8) + i)
                'index = index + 1
                'index = (((h - 1) * Resolution) + i)
                'index = IndexArray(h - 1) + i + 1

                index = index + 1

                'End While
            Next
        Next

            XPlacemark_Data.Save("C:\Google Earth Tour\PlacemarkData" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")
        XTrack.Save("C:\Google Earth Tour\Track" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")
        XAnimateModel.Save("C:\Google Earth Tour\Model" & Hour(Now) & Minute(Now) & Second(Now) & ".kml")

        'ProgressBar1.Value = ProgressBar1.Minimum

    End Sub
    '
    Private Function LoadPlacemarks(p1 As String) As Object

        'For pm = 0 To NPlacemarks
        '    XPlaceMark(pm) = XElement.Load(p1)
        'Next

        XPlaceMark(0) = XElement.Load(p1)

        Dim coordinates(5000) As String
        Dim TimeDataArray(5000) As String
        Dim HeadingDataArray(5000) As Double

        'If (XPlaceMark(0).Descendants(k + "Placemark").Count) > 0 Then
        'NPlacemarks = XPlaceMark(0).Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark").Count - 1
        NPlacemarks = XPlaceMark(0).Descendants(k + "Placemark").Count - 1

        'MsgBox(XPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)
        For pm = 0 To NPlacemarks

            coordinates(pm) = (XPlaceMark(0).Descendants(k + "Placemark")(pm).Elements(k + "Point").Elements(k + "coordinates").FirstOrDefault)

            If TimeInfoCheck.Checked = True And XPlaceMark(0).Descendants(k + "Placemark")(pm).Elements(k + "TimeSpan").Elements(k + "begin").Count > 0 Then
                TimeDataArray(pm) = (XPlaceMark(0).Descendants(k + "Placemark")(pm).Elements(k + "TimeSpan").Elements(k + "begin").FirstOrDefault)
            End If

            'If HeadingInfoCheck.Checked = True And XPlaceMark(0).Descendants(k + "Style").Elements(k + "IconStyle")(pm).Elements(k + "heading").Count > 0 Then
            If HeadingInfoCheck.Checked = True Then
                HeadingDataArray(pm) = CDbl(XPlaceMark(0).Descendants(k + "Style").Elements(k + "IconStyle")(pm).Elements(k + "heading").FirstOrDefault) Mod 360
            End If

            'MsgBox(coordinates(pm))
        Next


        Dim firstcomma(5000) As Integer

        For pm = 0 To NPlacemarks

            firstcomma(pm) = (InStr(coordinates(pm), ","))

        Next

        Dim longitudes(5000) As Double

        For pm = 0 To NPlacemarks
            longitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        For pm = 0 To NPlacemarks
            coordinates(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        For pm = 0 To NPlacemarks
            firstcomma(pm) = (InStr(coordinates(pm), ","))
        Next

        Dim latitudes(5000) As Double

        For pm = 0 To NPlacemarks
            latitudes(pm) = Microsoft.VisualBasic.Left(coordinates(pm), firstcomma(pm) - 1) * pi / 180
        Next

        Dim altitudes(5000) As Double

        For pm = 0 To NPlacemarks
            altitudes(pm) = Microsoft.VisualBasic.Right(coordinates(pm), Len(coordinates(pm)) - firstcomma(pm))
        Next

        Return {longitudes, latitudes, altitudes, TimeDataArray, HeadingDataArray}

    End Function
    Private Sub LoadModel()

        DaeNameSteps = -1

        For h = 0 To DaeModelListBox.Items.Count - 1
            DaeName(h) = DaeModelListBox.Items(h)
            DaeNameSteps = DaeNameSteps + 1
        Next
    End Sub



    Private Sub AddToPlacemarkData(HeadingString As String, BeginTime As String, EndTime As String, OutputString As String, OutputLongDeg As Double, OutputLatDeg As Double, altitudes As Double, OrientationString As String, TiltString As String, RangeString As String, index As Integer)


        Dim xAdd As XElement
        Dim xAddScreenOverlay As XElement
        Dim OverlayFilename As String


        Dim xInitialLookAt As XElement
        xInitialLookAt = <ns1:LookAt xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                             <ns2:TimeStamp>
                                 <ns1:when><%= BeginTime %></ns1:when>
                             </ns2:TimeStamp>
                             <ns1:longitude><%= OutputLongDeg %></ns1:longitude>
                             <ns1:latitude><%= OutputLatDeg %></ns1:latitude>
                             <ns1:altitude><%= altitudes %></ns1:altitude>
                             <ns1:heading><%= OrientationString %></ns1:heading>
                             <ns1:tilt><%= TiltString %></ns1:tilt>
                             <ns1:range><%= RangeString %></ns1:range>
                             <ns2:altitudeMode>relativeToGround</ns2:altitudeMode>
                         </ns1:LookAt>
        xAdd = <ns1:Placemark id="pm267" xmlns:ns1="http://www.opengis.net/kml/2.2">
                   <ns1:name><%= index %></ns1:name>
                   <ns1:Snippet maxLines="0">empty</ns1:Snippet>
                   <ns1:description>hello</ns1:description>
                   <ns1:TimeSpan>
                       <ns1:begin><%= BeginTime %></ns1:begin>
                       <ns1:end><%= EndTime %></ns1:end>
                   </ns1:TimeSpan>
                   <ns1:styleUrl>#Style_5</ns1:styleUrl>
                   <ns1:Point>
                       <ns1:altitudeMode>relativeToGround</ns1:altitudeMode>
                       <ns1:coordinates><%= OutputString %></ns1:coordinates>
                   </ns1:Point>
               </ns1:Placemark>

        OverlayFilename = "C:\Google Earth Tour\OverlayImages\MyImage" & index & ".png"

        xAddScreenOverlay = <ns1:ScreenOverlay xmlns:ns1="http://www.opengis.net/kml/2.2">
                                <ns1:name><%= BeginTime %></ns1:name>
                                <ns1:Icon>
                                    <ns1:href><%= OverlayFilename %></ns1:href>
                                </ns1:Icon>
                                <ns1:overlayXY x="0" y="-0.5" xunits="fraction" yunits="fraction"/>
                                <ns1:screenXY x="0" y="0" xunits="fraction" yunits="fraction"/>
                                <ns1:rotationXY x="0" y="0" xunits="fraction" yunits="fraction"/>
                                <ns1:size x="0" y="0" xunits="fraction" yunits="fraction"/>
                                <ns1:TimeSpan>
                                    <ns1:begin><%= BeginTime %></ns1:begin>
                                    <ns1:end><%= EndTime %></ns1:end>
                                </ns1:TimeSpan>
                            </ns1:ScreenOverlay>



        If index = 0 Then
            'XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index).ReplaceWith(xAdd)
            XPlacemark_Data.Elements(k + "Document").Elements(k + "LookAt")(index).ReplaceWith(xInitialLookAt)
            XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "ScreenOverlay")(index).ReplaceWith(xAddScreenOverlay)
        Else
            'XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index - 1).AddAfterSelf(xAdd)
            XPlacemark_Data.Elements(k + "Document").Elements(k + "Folder").Elements(k + "ScreenOverlay")(index - 1).AddAfterSelf(xAddScreenOverlay)
        End If

    End Sub

    Private Sub AddToAnimateModel(altitudeMode As String, horizFov As String, BeginTime As String, OutputLongDeg As Double, OutputLatDeg As Double, altitudes As Double, OrientationString As String, TiltString As String, RangeString As String, duration As Double, flyToMode As String, EndTime As String, ModelBearing As Double, DaeName As String, TrimData As Double, HeelData As Double, YawData As Double, index As Integer, ScaleData As Double)

        Dim xPlacemarkTable As XElement
        Dim xFlytoTable As XElement
        Dim xInitialLookAt As XElement

        Dim fixedlookat
        Dim fixedlongitude As Double
        Dim fixedlatitude As Double
        Dim fixedaltitude As Double
        Dim fixedheading As Double
        Dim fixedtilt As Double
        Dim fixedrange As Double

        If FixedLookAtCheckBox.Checked = True Then
            fixedlookat = ExtractLookAt(FixedLookAtTextBox.Text)
            fixedlongitude = fixedlookat(0)
            fixedlatitude = fixedlookat(1)
            fixedaltitude = fixedlookat(2)
            fixedheading = fixedlookat(3)
            fixedtilt = fixedlookat(4)
            fixedrange = fixedlookat(5)
        Else
            fixedlongitude = OutputLongDeg
            fixedlatitude = OutputLatDeg
            fixedaltitude = altitudes
            fixedheading = OrientationString
            fixedtilt = TiltString
            fixedrange = RangeString
        End If



        Dim ModelBearing2 As Double

        ModelBearing2 = ModelBearing + YawData

        xInitialLookAt = <ns1:LookAt xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                             <ns2:TimeStamp>
                                 <ns1:when><%= BeginTime %></ns1:when>
                             </ns2:TimeStamp>
                             <ns1:longitude><%= OutputLongDeg %></ns1:longitude>
                             <ns1:latitude><%= OutputLatDeg %></ns1:latitude>
                             <ns1:altitude><%= altitudes %></ns1:altitude>
                             <ns1:heading><%= OrientationString %></ns1:heading>
                             <ns1:tilt><%= TiltString %></ns1:tilt>
                             <ns1:range><%= RangeString %></ns1:range>
                             <ns2:altitudeMode>relativeToGround</ns2:altitudeMode>
                         </ns1:LookAt>


        xFlytoTable = <ns2:FlyTo xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                          <ns1:LookAt>
                              <ns2:altitudeMode><%= altitudeMode %></ns2:altitudeMode>
                              <ns2:horizFov><%= horizFov %></ns2:horizFov>
                              <ns2:TimeSpan>
                                  <ns1:begin><%= BeginTime %></ns1:begin>
                                  <ns1:end><%= BeginTime %></ns1:end>
                              </ns2:TimeSpan>
                              <ns1:longitude><%= fixedlongitude %></ns1:longitude>
                              <ns1:latitude><%= fixedlatitude %></ns1:latitude>
                              <ns1:altitude><%= fixedaltitude %></ns1:altitude>
                              <ns1:heading><%= fixedheading %></ns1:heading>
                              <ns1:tilt><%= fixedtilt %></ns1:tilt>
                              <ns1:range><%= fixedrange %></ns1:range>
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
                                          <ns1:heading><%= ModelBearing2 %></ns1:heading>
                                          <ns1:tilt><%= TrimData %></ns1:tilt>
                                          <ns1:roll><%= HeelData %></ns1:roll>
                                      </ns1:Orientation>
                                      <ns1:Scale>
                                          <ns1:x><%= ScaleData %></ns1:x>
                                          <ns1:y><%= ScaleData %></ns1:y>
                                          <ns1:z>1</ns1:z>
                                      </ns1:Scale>
                                      <ns1:Link>
                                          <ns1:href><%= DaeName %></ns1:href>
                                      </ns1:Link>
                                  </ns1:Model>
                              </ns1:MultiGeometry>
                          </ns1:Placemark>

        If index = 0 Then
            XAnimateModel.Elements(k + "Document").Elements(k + "LookAt")(index).ReplaceWith(xInitialLookAt)
            XAnimateModel.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index).ReplaceWith(xPlacemarkTable)
            XAnimateModel.Elements(k + "Document").Elements(kk + "Tour").Elements(kk + "Playlist").Elements(kk + "FlyTo")(index).ReplaceWith(xFlytoTable)

        Else
            XAnimateModel.Elements(k + "Document").Elements(k + "Folder").Elements(k + "Placemark")(index - 1).AddAfterSelf(xPlacemarkTable)
            XAnimateModel.Elements(k + "Document").Elements(kk + "Tour").Elements(kk + "Playlist").Elements(kk + "FlyTo")(index - 1).AddAfterSelf(xFlytoTable)

        End If


    End Sub



    Private Sub AddToTrack(BeginTime As String, CoordinateString As String, OutputLongDeg As Double, OutputLatDeg As Double, altitudes As Double, OrientationString As String, TiltString As String, RangeString As String, index As Integer)
        Dim xAddTime As XElement
        Dim xAddCoord As XElement
        Dim xInitialLookAt As XElement
        xInitialLookAt = <ns1:LookAt xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">
                             <ns2:TimeStamp>
                                 <ns1:when><%= BeginTime %></ns1:when>
                             </ns2:TimeStamp>
                             <ns1:longitude><%= OutputLongDeg %></ns1:longitude>
                             <ns1:latitude><%= OutputLatDeg %></ns1:latitude>
                             <ns1:altitude><%= altitudes %></ns1:altitude>
                             <ns1:heading><%= OrientationString %></ns1:heading>
                             <ns1:tilt><%= TiltString %></ns1:tilt>
                             <ns1:range><%= RangeString %></ns1:range>
                             <ns2:altitudeMode>relativeToGround</ns2:altitudeMode>
                         </ns1:LookAt>

        xAddTime = <ns1:when xmlns:ns1="http://www.opengis.net/kml/2.2"><%= BeginTime %></ns1:when>
        xAddCoord = <ns2:coord xmlns:ns2="http://www.google.com/kml/ext/2.2"><%= CoordinateString %></ns2:coord>

        If index = 0 Then
            XTrack.Elements(k + "Document").Elements(k + "LookAt")(index).ReplaceWith(xInitialLookAt)
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(k + "when")(index).ReplaceWith(xAddTime)
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(kk + "coord")(index).ReplaceWith(xAddCoord)
        Else
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(k + "when")(index - 1).AddAfterSelf(xAddTime)
            XTrack.Elements(k + "Document").Elements(k + "Placemark").Elements(kk + "Track").Elements(kk + "coord")(index - 1).AddAfterSelf(xAddCoord)
        End If

    End Sub



    Private Function MakeKML(TimeArray1 As String, TimeArray2 As String, LongLatAltString As String, HeadingArray As Double, n As Integer) As XElement

        Dim XElementAdd As XElement


        '<ns2:FlyTo xmlns:ns1="http://www.opengis.net/kml/2.2" xmlns:ns2="http://www.google.com/kml/ext/2.2">


        XElementAdd = <ns1:Document xmlns:ns1="http://www.opengis.net/kml/2.2">
                          <ns1:name><%= n %></ns1:name>
                          <ns1:open>1</ns1:open>

                          <ns1:StyleMap id="s_ylw-pushpin">
                              <ns1:Pair>
                                  <ns1:key>normal</ns1:key>
                                  <ns1:styleUrl>#s_ylw-pushpin1</ns1:styleUrl>
                              </ns1:Pair>
                              <ns1:Pair>
                                  <ns1:key>highlight</ns1:key>
                                  <ns1:styleUrl>#s_ylw-pushpin0</ns1:styleUrl>
                              </ns1:Pair>
                          </ns1:StyleMap>
                          <ns1:Style id="s_ylw-pushpin0">
                              <ns1:IconStyle>
                                  <ns1:scale>1.1</ns1:scale>
                                  <ns1:heading><%= HeadingArray %></ns1:heading>
                                  <ns1:Icon>
                                      <ns1:href>http://maps.google.com/mapfiles/kml/shapes/arrow.png</ns1:href>
                                  </ns1:Icon>
                                  <ns1:hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
                              </ns1:IconStyle>
                          </ns1:Style>
                          <ns1:Style id="s_ylw-pushpin1">
                              <ns1:IconStyle>
                                  <ns1:scale>1.1</ns1:scale>
                                  <ns1:heading><%= HeadingArray %></ns1:heading>
                                  <ns1:Icon>
                                      <ns1:href>http://maps.google.com/mapfiles/kml/shapes/arrow.png</ns1:href>
                                  </ns1:Icon>
                                  <ns1:hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
                              </ns1:IconStyle>
                          </ns1:Style>
                          <ns1:Placemark>
                              <ns1:name><%= n %></ns1:name>
                              <ns1:open>1</ns1:open>
                              <ns1:styleUrl>#s_ylw-pushpin</ns1:styleUrl>
                              <ns1:TimeSpan>
                                  <ns1:begin><%= TimeArray1 %></ns1:begin>
                                  <ns1:end><%= TimeArray2 %></ns1:end>
                              </ns1:TimeSpan>
                              <ns1:Point>
                                  <ns1:coordinates><%= LongLatAltString %></ns1:coordinates>
                              </ns1:Point>
                          </ns1:Placemark>
                      </ns1:Document>


        Return XElementAdd
    End Function

    ' min and max are the minimum and maximum indexes of the items that might still be out of order.
    Sub BubbleSort(List As Double(,), ByVal min As Integer, ByVal max As Integer)
        Dim last_swap As Integer
        Dim m As Integer
        Dim n As Integer
        Dim tmp(3) As Double

        ' Repeat until we are done.
        Do While min < max
            ' Bubble up.
            last_swap = min - 1
            ' For i = min + 1 To max
            m = min + 1
            Do While m <= max
                ' Find a bubble.
                If List(m - 1, 0) > List(m, 0) Then
                    ' See where to drop the bubble.
                    For h = 0 To 3
                        tmp(h) = List(m - 1, h)
                    Next
                    n = m
                    Do
                        For h = 0 To 3
                            List(n - 1, h) = List(n, h)
                        Next
                        n = n + 1
                        If n > max Then Exit Do
                    Loop While List(n, 0) < tmp(0)
                    For h = 0 To 3
                        List(n - 1, h) = tmp(h)
                    Next
                    last_swap = n - 1
                    m = n + 1
                Else
                    m = m + 1
                End If
            Loop
            ' Update max.
            max = last_swap - 1

            ' Bubble down.
            last_swap = max + 1
            ' For i = max - 1 To min Step -1
            m = max - 1

            Do While m >= min

                ' Find a bubble.
                If List(m + 1, 0) < List(m, 0) Then
                    'MsgBox(m)
                    For h = 0 To 3
                        ' See where to drop the bubble.
                        tmp(h) = List(m + 1, h)
                    Next
                    n = m
                    Do
                        For h = 0 To 3
                            List(n + 1, h) = List(n, h)
                        Next
                        n = n - 1
                        If n < min Then Exit Do
                    Loop While List(n, 0) > tmp(0)
                    For h = 0 To 3
                        List(n + 1, h) = tmp(h)
                    Next
                    last_swap = n + 1
                    m = n - 1

                Else
                    m = m - 1
                End If
            Loop
            ' Update min.
            min = last_swap + 1
        Loop
    End Sub


    Private Function h00(t As Double) As Double
        Return (2 * t ^ 3 - 3 * t ^ 2 + 1)
    End Function


    Private Function kmlDate(m4 As Date) As String
        'MsgBox(FormatMilliseconds(m4))
        'Return Year(m4) & "-" & Format(Month(m4), "00") & "-" & Format(Day(m4), "00") & "T" & Format(Hour(m4), "00") & ":" & Format(Minute(m4), "00") & ":" & Format(Second(m4), "00") & "." & m4.Millisecond & "Z"
        Return Year(m4) & "-" & Format(Month(m4), "00") & "-" & Format(Day(m4), "00") & "T" & Format(Hour(m4), "00") & ":" & Format(Minute(m4), "00") & ":" & Format(Second(m4), "00") & "." & FormatMilliseconds(m4) & "Z"

    End Function

    Function FormatMilliseconds(timeValue As Date, Optional formatString As String = "000") As String
        Dim milliseconds As Integer
        'milliseconds = (Hour(timeValue) * 3600 + Minute(timeValue) * 60 + Second(timeValue)) * 1000 + timeValue.Millisecond
        milliseconds = timeValue.Millisecond

        FormatMilliseconds = Format(milliseconds, formatString)
    End Function


    '    Private Sub CreateImageReadout(BeginTime As String, VeryEndTime As String, OutputLongDeg As Double, OutputLatDeg As Double, OrientationString As String, TiltString As String, RangeString As String, indexForPlacemarkData As Integer, TrimString As String, HeelString As String, HeadingString As String, SpeedString As String, DraftString As String, YawString As String)
    Private Sub CreateImageReadout(BeginTime As String, VeryEndTime As String, OutputLongDeg As Double, OutputLatDeg As Double, OrientationString As String, TiltString As String, RangeString As String, indexForPlacemarkData As Integer, PassengerString As String, WeatherString As String, HeadingString As String, SpeedString As String, PowerString As String, BatteryString As String, ReadoutTitleString As String, ReadoutSubTitleString As String)
        'MsgBox(BeginTime)
        Dim FontColor As Color = Color.White
        Dim BackColor As Color = Color.FromArgb(128, Color.Black) 'transparent
        Dim FontName As String = "montserrat" 'courier
        Dim FontSize As Integer = 12
        Dim Height As Integer = 270
        Dim Width As Integer = 400
        Dim objBitmap As New Bitmap(Width, Height)

        'Dim myBitmap As New Bitmap("C:\Google Earth Tour\callum logo transparent background.png")
#Disable Warning BC42024 ' Unused local variable
        Dim myBitmap As Bitmap '("C:\Google Earth Tour\callum logo transparent background.png")
#Enable Warning BC42024 ' Unused local variable


        Dim FileName As String = "MyImage"
        Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)
        Dim objFont As New Font(FontName, FontSize)
        Dim objFontTitle As New Font(FontName, 18, FontStyle.Bold)

        Dim obj0 As New PointF(10.0F, 20.0F)
        Dim obj01 As New PointF(10.0F, 20.0F)
        Dim obj1 As New PointF(10.0F, 20.0F)
        Dim obj2 As New PointF(10.0F, 20.0F)
        Dim obj3 As New PointF(10.0F, 20.0F)
        Dim obj4 As New PointF(10.0F, 20.0F)
        Dim obj5 As New PointF(10.0F, 20.0F)
        Dim obj6 As New PointF(10.0F, 20.0F)
        Dim obj7 As New PointF(10.0F, 20.0F)
        Dim obj8 As New PointF(10.0F, 20.0F)
        Dim obj9 As New PointF(10.0F, 20.0F)
        Dim obj10 As New PointF(10.0F, 20.0F)



        Dim objBrushForeColor As New SolidBrush(FontColor)

        Dim objBrushForeColor1 As New SolidBrush(Color.White) 'DarkOrange

        Dim objBrushBackColor As New SolidBrush(BackColor)


        'objBitmap.MakeTransparent()
        ''objGraphics.FillRectangle(objBrushBackColor, 0, 0, Width, Height)
        'objGraphics.DrawString("Time: " & BeginTime, objFont, objBrushForeColor, objPoint)
        'objGraphics.DrawString("Lat: " & OutputLatDeg, objFont, objBrushForeColor, objPoint1)

        'objBitmap.Save("C:\Google Earth Tour\" & FileName & indexForPlacemarkData.ToString & ".bmp", ImageFormat.Bmp)

        ' Create a Bitmap object from an image file. 

        Dim myRectangle As New Rectangle
        Dim myRectangle0 As New Rectangle
        Dim myRectangle1 As New Rectangle
        Dim myRectangle2 As New Rectangle
        Dim myRectangle3 As New Rectangle
        Dim myRectangle4 As New Rectangle
        Dim myRectangle5 As New Rectangle
        Dim myRectangle6 As New Rectangle
        Dim myRectangle7 As New Rectangle
        Dim myRectangle8 As New Rectangle
        Dim myRectangle9 As New Rectangle





        ' Draw myBitmap to the screen.

        obj0.X = 5
        obj01.X = 5
        obj1.X = 15
        obj2.X = 15
        obj3.X = 15
        obj4.X = 15
        obj5.X = 15
        obj6.X = 15
        obj7.X = 15
        obj8.X = 15
        obj9.X = 15
        obj10.X = 15


        Dim constpix As Integer = 30

        obj0.Y = 0  'myBitmap.Height + 20
        obj01.Y = 15  'myBitmap.Height + 20
        obj1.Y = 25 + constpix 'myBitmap.Height + 20
        obj2.Y = 45 + constpix ' myBitmap.Height + 50
        obj3.Y = 65 + constpix 'myBitmap.Height + 80
        obj4.Y = 85 + constpix 'myBitmap.Height + 110
        obj5.Y = 105 + constpix 'myBitmap.Height + 110
        obj6.Y = 125 + constpix 'myBitmap.Height + 110
        obj7.Y = 145 + constpix 'myBitmap.Height + 110
        obj8.Y = 165 + constpix 'myBitmap.Height + 110
        obj9.Y = 185 + constpix  'myBitmap.Height + 110
        obj10.Y = 205 + constpix  'myBitmap.Height + 110


        myRectangle.X = 10
        myRectangle.Y = 10 + constpix ' myBitmap.Height + 10
        myRectangle.Height = 220
        myRectangle.Width = 320 ' myBitmap.Width - 15

        myRectangle1.X = 10
        myRectangle1.Y = constpix ' myBitmap.Height + 10
        myRectangle1.Height = 20
        myRectangle1.Width = 80 ' myBitmap.Width - 15

        myRectangle2.X = 10
        myRectangle2.Y = 20 + constpix ' myBitmap.Height + 10
        myRectangle2.Height = 20
        myRectangle2.Width = 80 ' myBitmap.Width - 15

        myRectangle3.X = 10
        myRectangle3.Y = 40 + constpix ' myBitmap.Height + 10
        myRectangle3.Height = 20
        myRectangle3.Width = 80 ' myBitmap.Width - 15

        'objGraphics.DrawImage(myBitmap, 0, 0, myBitmap.Width, myBitmap.Height)

        Dim OrangePen As New Pen(Color.OrangeRed, 5)
        OrangePen.Alignment = Drawing2D.PenAlignment.Center


        objGraphics.DrawString(ReadoutSubTitleString, objFont, objBrushForeColor1, obj0)

        'objGraphics.DrawString("GIBSONS → VANCOUVER", objFontTitle, objBrushForeColor1, obj01)
        objGraphics.DrawString(ReadoutTitleString, objFontTitle, objBrushForeColor1, obj01)


        'objGraphics.DrawRectangle(OrangePen, myRectangle)
        objGraphics.FillRectangle(objBrushBackColor, myRectangle)


        'objGraphics.DrawRectangle(OrangePen, myRectangle1)
        'objGraphics.FillRectangle(Brushes.DarkOrange, myRectangle1)
        'objGraphics.DrawRectangle(OrangePen, myRectangle2)
        'objGraphics.FillRectangle(Brushes.OrangeRed, myRectangle2)
        'objGraphics.DrawRectangle(OrangePen, myRectangle3)
        'objGraphics.FillRectangle(Brushes.Orange, myRectangle2)







        objGraphics.DrawString("Date: " & Split(BeginTime, "T")(0), objFont, objBrushForeColor, obj1)
        'MsgBox(("Time: " & Split(BeginTime, "T")(1)))
        objGraphics.DrawString("Time: " & Split(BeginTime, "T")(1), objFont, objBrushForeColor, obj2)
        objGraphics.DrawString("Latitude: " & OutputDegString(OutputLatDeg), objFont, objBrushForeColor, obj3)
        objGraphics.DrawString("Longitude: " & OutputDegString(OutputLongDeg), objFont, objBrushForeColor, obj4)
        objGraphics.DrawString(HeadingString, objFont, objBrushForeColor, obj5)
        objGraphics.DrawString(SpeedString, objFont, objBrushForeColor, obj6)


        objGraphics.DrawString(WeatherString, objFont, objBrushForeColor, obj7)
        objGraphics.DrawString(PassengerString, objFont, objBrushForeColor, obj8)
        objGraphics.DrawString(PowerString, objFont, objBrushForeColor, obj9)
        objGraphics.DrawString(BatteryString, objFont, objBrushForeColor, obj10)






        ' Make the default transparent color transparent for myBitmap.
        'myBitmap.MakeTransparent()

        ' Draw the transparent bitmap to the screen.
        'objGraphics.DrawImage(myBitmap, myBitmap.Width, 0, myBitmap.Width, myBitmap.Height)


        If ReadoutCheckBox.Checked = True Then
            objBitmap.Save("C:\Google Earth Tour\OverlayImages\" & FileName & indexForPlacemarkData.ToString & ".png", ImageFormat.Png)
        End If

    End Sub

    Private Function OutputDegString(OutputLatDeg As Double) As String
        'Return Int(OutputLatDeg) & "°" & Int((OutputLatDeg - Int(OutputLatDeg)) / 24 * (24 * 60) Mod 60) & "'" & Math.Round((OutputLatDeg - Int(OutputLatDeg)) / 24 * (24 * 60 * 60) Mod 60, 2) & "'';"
        Return OutputLatDeg.ToString
    End Function

    Private Function ExtractLookAt(p1 As String) As Object

        Dim Longitude As Double
        Dim Latitude As Double
        Dim Altitude As Double
        Dim Heading As Double
        Dim Tilt As Double
        Dim Range As Double


        XLookAtPlaceMark(0) = XElement.Load(p1)

        Longitude = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "longitude").FirstOrDefault)
        Latitude = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "latitude").FirstOrDefault)
        Altitude = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "altitude").FirstOrDefault)
        Heading = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "heading").FirstOrDefault)
        Tilt = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "tilt").FirstOrDefault)
        Range = (XLookAtPlaceMark(0).Descendants(k + "Placemark")(0).Elements(k + "LookAt").Elements(k + "range").FirstOrDefault)

        'MsgBox(Range)

        Return {Longitude, Latitude, Altitude, Heading, Tilt, Range}
    End Function


End Class
