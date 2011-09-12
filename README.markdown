## inConvexPolygon ##
`inConvexPolygon` takes a point (_PointX_, _PointY_) along with four other points (_topMostX_, _topMostY_), (_rightMostX_, _rightMostY_), (_bottomMostX_, _bottomMostY_), and (_leftMostX_, _leftMostY_) that define a convex polygon and returns `true` if the passed point is within the polygon's area, `false` if not.

A picture is worth a thousand words:

![Points within convex polygon](https://lh4.googleusercontent.com/-PZT7LOt01UI/Tm1IckHofhI/AAAAAAAAAaI/s7qLunE0BMY/points-in-polygon.png)

_red points found by using `inConvexPolygon` function - sample app can be downloaded at sample/in\_convex\_polygon.xlsm_

A quick note on polygon points: it is possible that a polygon can have a point that can be both the left most and the top most (or the left most and the bottom most, etc.).
In a situation such as this, just remember to pick a starting point and then assign the remaining points based on clockwise traversal of the polygon's sides.

## Basic Usage ##
Import __modules/convex\_polygon.bas__ module into any new or existing MS Excel workbook. You can then use `inConvexPolygon` to determine if any point is within an arbitrary 4-sided convex polygon.

The function uses the _"to the right rule"_ to determine if the point is inside the polygon's bounded area. A point is in the polygon's bounded area if the point is to the right of all the line segments when traversed clockwise. Stated more simply: it compares the slope of the given point to the starting point of the given line segment with the slope of the given line segment.

## Todo ##
_Please feel free push any code that implements the following - patches will be happily
accepted_

* Better naming for topMostX, topMostY, etc. as a single point could be both;
  current naming could be viewed as confusing in a situation like this.
* Generalize #inConvexPolygon to handle an n-sided convex polygon; current
  implementation handles a 4-sided convex polygon only (this was my use case)

Any others come to mind? Email me at [jeff.iacono@gmail.com](mailto:jeff.iacono@gmail.com).

## Further Reading & Examples ##
See: http://paulbourke.net/geometry/insidepoly/
