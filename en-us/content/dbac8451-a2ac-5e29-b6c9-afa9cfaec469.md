
# CapitalizeNamesOfDays Property

 **Last modified:** June 30, 2011

 _**Applies to:** Excel 2013 | Office 2013 | VBA_

True if the first letter of day names is capitalized automatically. Read/write Boolean.

 _expression_. **CapitalizeNamesOfDays**
 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets Microsoft Graph to capitalize the first letter of the names of days.


```
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```


****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/dbac8451-a2ac-5e29-b6c9-afa9cfaec469.md) using GitHub.

