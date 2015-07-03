
# PivotField.SubtotalName Property (Excel)

Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string "Subtotal". Read/write  **String**.


## Syntax

 _expression_. **SubtotalName**

 _expression_A variable that represents a  **PivotField** object.


## Example

This example sets the subtotal label to "Regional Subtotal" (instead of the default string "Subtotal") in the state field in the second PivotTable report on the active worksheet.


```
ActiveSheet.PivotTables("PivotTable2") _ 
 .PivotFields("state").SubtotalName = "Regional Subtotal"
```


## See also


#### Concepts


 [PivotField Object](52784960-e2da-b43a-1e37-2d4dae61c6d8.md)
#### Other resources


 [PivotField Object Members](4a6ea12a-072c-a386-c855-7bf5f6eadd46.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/db2f8366-75a4-edca-f46f-f0bff083ccbe.md) using GitHub.

