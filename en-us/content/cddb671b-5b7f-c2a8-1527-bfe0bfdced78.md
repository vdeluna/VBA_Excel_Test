
# Window.DisplayZeros Property (Excel)

 **True** if zero values are displayed. Read/write **Boolean**.


## Syntax

 _expression_. **DisplayZeros**

 _expression_A variable that represents a  **Window** object.


## Remarks

This property applies only to worksheets and macro sheets.


## Example

This example sets the active window in Book1.xls to display zero values.


```
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.DisplayZeros = True 

```


## See also


#### Concepts


 [Window Object](8591b1ad-76f8-14e2-9120-406b65093f5a.md)
#### Other resources


 [Window Object Members](f11db427-24a4-041c-2fd5-03ce73ae6c16.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/cddb671b-5b7f-c2a8-1527-bfe0bfdced78.md) using GitHub.

