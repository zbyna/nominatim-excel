Attribute VB_Name = "Module1"
Option Explicit


Public Function ReverseGeocode(lat As String, lng As String) As String

' https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=-34.44076&lon=-58.70521

   Dim nominatimClient As New WebClient
    nominatimClient.BaseUrl = "https://nominatim.openstreetmap.org/reverse"

    ' Create a WebRequest for getting info
    Dim nominatimRequest As New WebRequest
    nominatimRequest.Method = WebMethod.HttpGet
    nominatimRequest.ResponseFormat = WebFormat.json

    ' Set the request format
    nominatimRequest.Format = WebFormat.json

    ' Add querystring to the request
    nominatimRequest.AddQuerystringParam "format", "jsonv2"
    nominatimRequest.AddQuerystringParam "lat", lat
    nominatimRequest.AddQuerystringParam "lon", lng
    ' Add custom header
    nominatimRequest.SetHeader "Accept-Language", "cs"

    ' Execute the request and work with the response
    Dim Response As WebResponse
    Set Response = nominatimClient.Execute(nominatimRequest)

    'Debug.Print Response.StatusCode
    'Debug.Print Response.Content
  
  If Response.StatusCode = 200 Then
    If Response.Data("error") <> "" Then
         ReverseGeocode = Response.Data("error")
    Else
        ReverseGeocode = Response.Data("display_name")
    End If
  Else
    ReverseGeocode = Response.Data("error")("message")
  End If
  
  ' release memory
  Set nominatimClient = Nothing
  Set nominatimRequest = Nothing
  Set Response = Nothing
  'Application.Calculation = xlAutomatic
End Function




Public Function Geocode(adressToSearch As String) As String

' https://nominatim.openstreetmap.org/search?
'         format=json&addressdetails=1&q=bakery+in+berlin+wedding&format=json&limit=1

   Dim nominatimClient As New WebClient
    nominatimClient.BaseUrl = "https://nominatim.openstreetmap.org/search"

    ' Create a WebRequest for getting info
    Dim nominatimRequest As New WebRequest
    nominatimRequest.Method = WebMethod.HttpGet
    nominatimRequest.ResponseFormat = WebFormat.json

    ' Set the request format
    nominatimRequest.Format = WebFormat.json

    ' Add querystring to the request
    nominatimRequest.AddQuerystringParam "format", "json"
    nominatimRequest.AddQuerystringParam "limit", 1
    
    Debug.Print "adressToSearch: " & adressToSearch
    
    nominatimRequest.AddQuerystringParam "q", adressToSearch
    
    ' Add custom header
    nominatimRequest.SetHeader "Accept-Language", "cs"
    
    Debug.Print nominatimRequest.FormattedResource

    ' Execute the request and work with the response
    Dim Response As WebResponse
    Set Response = nominatimClient.Execute(nominatimRequest)

    Debug.Print Response.StatusCode
    Debug.Print Response.Content
  
  If Response.StatusCode = 200 Then
        Geocode = Response.Data(1)("lat") + "," + Response.Data(1)("lon")
        
        Debug.Print "Geocode: " & Geocode
        
  Else
    Geocode = Response.Data("error")("message")
  End If
  
  ' release memory
  Set nominatimClient = Nothing
  Set nominatimRequest = Nothing
  Set Response = Nothing
  'Application.Calculation = xlAutomatic
End Function



