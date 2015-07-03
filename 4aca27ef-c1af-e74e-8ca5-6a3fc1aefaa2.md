
# MinimumScale Property

 **Last modified:** June 30, 2011

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

Returns or sets the minimum value on the axis. Read/write  **Double**.


## Remarks

Setting this property sets the  ** [MinimumScaleIsAuto](95ed7a2b-efda-b05a-da2e-789a166a97c8.md)**property to  **False**.


## Example

This example sets the minimum and maximum values for the value axis.


```
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/4aca27ef-c1af-e74e-8ca5-6a3fc1aefaa2.md) using GitHub.

