
# CustomProperty.Application Property (Excel)

When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **CustomProperty** object.


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


 [CustomProperty Object](df8b58d8-ccfd-00bb-723a-a9c328f0b38b.md)
#### Other resources


 [CustomProperty Object Members](a63c6fa9-2a9f-745a-987c-f977bf2c679a.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/c62cc90e-f672-01be-da63-0cdb842adbec.md) using GitHub.

