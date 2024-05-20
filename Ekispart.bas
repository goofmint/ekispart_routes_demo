Attribute VB_Name = "Ekispart"
Type Station
    Name As String
    Code As String
End Type

Type Line
    Name As String
    Type As String
    End Type
    
Type Route
    Line As Line
    Boarding As Station
    Destination As Station
End Type

Type Course
    Price As Long
    Route() As Route
End Type
    
    
