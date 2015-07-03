
# HasLegend Property

 **Last modified:** June 30, 2011

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

 **True** if the chart has a legend. Read/write **Boolean**.


## Example

This example turns on the legend for the chart and then sets the legend font color to blue.


```
With myChart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
End With
```


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/b4dbef39-9d83-2f6e-fe06-8ca38cceeeec.md) using GitHub.

