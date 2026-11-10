Sub zapInk()
    Dim osld As Slide
    Dim L As Long
    For Each osld In ActivePresentation.Slides
        For L = osld.Shapes.Count To 1 Step -1
            If osld.Shapes(L).Type = 23 Then osld.Shapes(L).Delete
        Next L
    Next osld
End Sub
