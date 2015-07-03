
# ListRow.Application Property (Excel)

When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **ListRow** object.


## Example

This example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## See also


#### Concepts


 [ListRow Object](ba3e4215-14b6-3dca-82d0-0951f9f2fc3e.md)
#### Other resources


 [ListRow Object Members](cd5e2170-7193-d865-f9f4-ce247e27c2f9.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/ace7b6cb-4afe-d2d5-5912-e2eb1aa5a90a.md) using GitHub.

