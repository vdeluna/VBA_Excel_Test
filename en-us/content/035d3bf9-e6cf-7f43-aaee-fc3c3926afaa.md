
# MarkerBackgroundColor Property

 **Last modified:** June 30, 2011

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

Returns or sets the marker background color as an RGB value. Applies only to line, scatter, and radar charts. Read/write  **Long**.


## Example

This example sets the marker background and foreground colors for the second point in series one.


```
With myChart.SeriesCollection(1).Points(2) 
 .MarkerBackgroundColor = RGB(0,255,0) ' green 
 .MarkerForegroundColor = RGB(255,0,0) ' red 
End With
```


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/035d3bf9-e6cf-7f43-aaee-fc3c3926afaa.md) using GitHub.

