
# AxisTitle.Text Property (Excel)

Returns or sets the text for the specified object. Read/write  **String**.


## Syntax

 _expression_. **Text**

 _expression_A variable that represents an  **AxisTitle** object.


## Example

This example sets the axis title text for the category axis in Chart1.


```
With Charts("Chart1").Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
End With
```


## See also


#### Concepts


 [AxisTitle Object](563d3ba5-aa77-b6fc-236a-7838d75eaa53.md)
#### Other resources


 [AxisTitle Object Members](84970b5a-91a1-b785-5632-97a0de4410f2.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/1305fae5-afd9-dd8e-f559-f0c6ebff7a3b.md) using GitHub.

