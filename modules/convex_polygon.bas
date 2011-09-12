Attribute VB_Name = "convex_polygon"
Option Explicit

' Usage: returns true if the point (pointX, pointY) is to the right of the line segement
'        defined by the starting point (startX, startY) and ending point (endX, endY).
'        This can be re-written to demonstrate that it is checking that the slope of the line
'        from the point to the starting point is less than or equal to the slope of the line from end point
'        to the start point:
'        (delta(y, point => start)) / (delta(x, point => start)) <= (delta(y, end => start)) / (delta(x, end => start))
Private Function isToTheRight(pointY As Double, pointX As Double, _
                                startY As Double, startX As Double, _
                                endY As Double, endX As Double) As Boolean
    isToTheRight = (((pointY - startY) * (endX - startX) - (pointX - startX) * (endY - startY)) <= 0)
End Function

' Usage: Determine if the point (pointX, pointY) is in the convex polygon created by the points
' (topMostX, topMostY), (rightMostX, rightMostY), (bottomMostX, bottomMostY), and
' (leftMostX, leftMostY) by using the "to the right rule": a point is in the polygon's bounded area
' if the point is to the right of the different line segments when traversed clockwise
'
' In pictures:
'
'                topMost (x,y)
'                     p
'                    / \
'                   /   \
'  leftMost (x,y)  p     p  rightMost (x,y)
'                   \   /
'                    \ /
'                     p
'              bottomMost (x,y)
'
' Note: assumes a 4 sided polygon, can be expanded to handle n-sided convex polygon, left as future dev project
' For further discussion / examples, see: http://paulbourke.net/geometry/insidepoly/
Public Function inConvexPolygon(pointX As Double, pointY As Double, _
                                    topMostX As Double, topMostY As Double, _
                                    rightMostX As Double, rightMostY As Double, _
                                    bottomMostX As Double, bottomMostY As Double, _
                                    leftMostX As Double, leftMostY As Double) As Boolean
                                
    ' Return false if we encounter a failing scenario, no sense in running any remaining tests
    ' Note: we *must* use clunky if-blocks as VBA does *not* short circuit its if condition statements
    '       when a false is found in AND'd conditions.
    '       For Example:
    '
    '           Public Function somethingThatShouldnotBeCalled() As Boolean
    '               Debug.Print "I should not be called"
    '               somethingThatShouldnotBeCalled = False
    '           End Function
    '
    '           Public Function testShortCircuit() As Boolean
    '               If False And somethingThatShouldnotBeCalled() Then
    '                   Debug.Print "I should never print"
    '               End If
    '               testShortCircuit = True ' for completeness' sake
    '           End Function
    '
    '           ' (run in immediate window w/ ctrl + g)
    '           testShortCircuit() '=> I should not be called / True
    '
    '       Because we have an falsed AND conditional statement, one would expect the if to short
    '       circuit and not run the other test, however this is not the case. Sad pig.
    
    If Not isToTheRight(pointY, pointX, topMostY, topMostX, rightMostY, rightMostX) Then
        inConvexPolygon = False: Exit Function
    End If
    
    If Not isToTheRight(pointY, pointX, rightMostY, rightMostX, bottomMostY, bottomMostX) Then
        inConvexPolygon = False: Exit Function
    End If
    
    If Not isToTheRight(pointY, pointX, bottomMostY, bottomMostX, leftMostY, leftMostX) Then
        inConvexPolygon = False: Exit Function
    End If
    
    If Not isToTheRight(pointY, pointX, leftMostY, leftMostX, topMostY, topMostX) Then
        inConvexPolygon = False: Exit Function
    End If

    inConvexPolygon = True
End Function
