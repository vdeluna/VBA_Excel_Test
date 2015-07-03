
# ChartGroup.BubbleScale Property (Excel)

Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. Read/write  **Long**.


## Syntax

 _expression_. **BubbleScale**

 _expression_A variable that represents a  **ChartGroup** object.


## Example

This example sets the bubble size in chart group one to 200% of the default size.


```
With Worksheets(1).ChartObjects(1).Chart 
 .ChartGroups(1).BubbleScale = 200 
End With
```


## See also


#### Concepts


 [ChartGroup Object](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)
#### Other resources


 [ChartGroup Object Members](2d31f7af-d639-c8f4-0714-08fc618ec92d.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/cbab742e-4e60-2d10-e8ec-0dcd2a5ff72a.md) using GitHub.

