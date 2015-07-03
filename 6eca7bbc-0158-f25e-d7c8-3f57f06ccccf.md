
# ChartTitle Object

 **Last modified:** June 30, 2011

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

Represents the title of the specified chart.


## Using the ChartTitle Object

Use the  **ChartTitle** property to return the **ChartTitle** object. The following example adds a title to the chart.


```
With myChart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## Remarks

The  **ChartTitle** object doesn't exist and cannot be used unless the ** [HasTitle](9ecc48d3-fd86-e185-a69f-0676230b3194.md)**property for the chart is  **True**.


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/6eca7bbc-0158-f25e-d7c8-3f57f06ccccf.md) using GitHub.

