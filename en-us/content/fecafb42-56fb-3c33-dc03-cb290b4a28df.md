
# Chart.AutoScaling Property (Excel)

 **True** if Microsoft Excel scales a 3-D chart so that it's closer in size to the equivalent 2-D chart. The ** [RightAngleAxes](632aa454-4113-97d3-a80c-eb745a950c6f.md)**property must be  **True**. Read/write  **Boolean**.


## Syntax

 _expression_. **AutoScaling**

 _expression_A variable that represents a  **Chart** object.


## Example

This example automatically scales Chart1. The example should be run on a 3-D chart.


```
With Charts("Chart1") 
 .RightAngleAxes = True 
 .AutoScaling = True 
End With
```


## See also


#### Concepts


 [Chart Object](179c32ce-49bd-6f36-ea12-89fb5443f3ea.md)
#### Other resources


 [Chart Object Members](a3f8ac44-02d6-6f3f-b5e0-23f4bd5d6baf.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/fecafb42-56fb-3c33-dc03-cb290b4a28df.md) using GitHub.

