Attribute VB_Name = "Module1"
Option Explicit

Sub EmptyColl()

    Dim coll As New Collection
    
    Set coll = Nothing
    
    coll.Add "Pear"
    
End Sub
