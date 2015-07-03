
# Worksheet.Comments Property (Excel)

Returns a  ** [Comments](f43bf021-1e46-10cf-09bf-070fc6a2c81a.md)**collection that represents all the comments for the specified worksheet. Read-only.


## Syntax

 _expression_. **Comments**

 _expression_A variable that represents a  **Worksheet** object.


## Example

This example deletes all comments added by Jean Selva on the active sheet.


```
For Each c in ActiveSheet.Comments 
 If c.Author = "Jean Selva" Then c.Delete 
Next
```


## See also


#### Concepts


 [Worksheet Object](182b705e-854a-81cc-a4b0-59b942de55ae.md)
#### Other resources


 [Worksheet Object Members](f8c1afea-1a1c-f5e4-37e3-52c434c8c157.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/c2ad8ea7-0fa3-7cde-e3f2-49bbdb81d261.md) using GitHub.

