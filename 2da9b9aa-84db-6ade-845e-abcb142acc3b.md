
# BarShape Property

 **Last modified:** March 09, 2015

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

Returns or sets the shape used with the specified 3-D bar or column chart. Read/write XlBarShape .



|XlBarShape can be one of these XlBarShape constants.|
| **xlConeToMax**|
| **xlCylinder**|
| **xlPyramidToPoint**|
| **xlBox**|
| **xlConeToPoint**|
| **xlPyramidToMax**|
 _expression_. **BarShape**
 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the shape used with series one on the chart.


```
myChart.SeriesCollection(1).BarShape = xlConeToPoint
```


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/2da9b9aa-84db-6ade-845e-abcb142acc3b.md) using GitHub.

