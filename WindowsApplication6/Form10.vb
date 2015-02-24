'http://www.c-sharpcorner.com/uploadfile/scottlysle/geocode-an-address-using-google-maps-in-vb-net/

Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Collections.Generic
Imports System.ComponentModel
Public Class Form10
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    End Sub
    Private Sub btnGeocode_Click(sender As System.Object, e As System.EventArgs) Handles btnGeocode.Click
        Try
            txtLatLon.Text = GetLatLon(txtAddress.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "An error has occurred")
        End Try
    End Sub

    Public Function GetLatLon(ByVal addr As String) As String
        Dim googleResult As New MSXML2.DOMDocument
        Dim googleService As New MSXML2.XMLHTTP
        Dim oNodes As MSXML2.IXMLDOMNodeList
        Dim oNode As MSXML2.IXMLDOMNode

        Dim strLatitude As String
        Dim strLongitude As String

        addr = URLEncode(addr)

        Dim url As String = "https://maps.googleapis.com/maps/api/geocode/xml?"
        url = url & "address=" & addr
        url = url & "&key=AIzaSyBIIiqxDh5w2De_B7f2YQH_QsnwQnMrPXg"

        googleService.open("GET", url, False)
        googleService.send()
        googleResult.loadXML(googleService.responseText)

        oNodes = googleResult.getElementsByTagName("geometry")

        If oNodes.length = 1 Then
            For Each oNode In oNodes
                strLatitude = oNode.childNodes(0).childNodes(0).text
                strLongitude = oNode.childNodes(0).childNodes(1).text
                GetLatLon = strLatitude & "," & strLongitude
            Next oNode
        Else
            GetLatLon = "Not Found (try again, you may have done too many too fast)"

        End If


    End Function
    Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
        Dim CharCode As Integer
        Dim Char1 As String
        Dim Space As String
        Dim StringLen As Long
        Dim i As Long
        StringLen = Len(StringVal)
        Dim result(StringLen) As String
        If StringLen > 0 Then

            If SpaceAsPlus Then Space = "+" Else Space = "%20"

            For i = 1 To StringLen
                Char1 = Mid$(StringVal, i, 1)

                CharCode = Asc(Char1)
                Select Case CharCode
                    Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                        result(i) = Char1
                    Case 32
                        result(i) = Space
                    Case 0 To 15
                        result(i) = "%0" & Hex(CharCode)
                    Case Else
                        result(i) = "%" & Hex(CharCode)
                End Select
            Next i
            URLEncode = Join(result, "")
        End If
    End Function


End Class

