
# HPageBreak.Application Property (Excel)

When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **HPageBreak** object.


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


 [HPageBreak Object](8fc96958-33ab-8251-f627-4769b5eab97f.md)
#### Other resources


 [HPageBreak Object Members](32b561ff-a0cf-142b-0a46-c622a42b6125.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/d0a1b77f-f5b9-7cdf-cbf4-25b59d86ace7.md) using GitHub.

