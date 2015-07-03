
# XmlDataBinding.Application Property (Excel)

When used without an object qualifier, this property returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)**object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an  **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  **XmlDataBinding** object.


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


 [XmlDataBinding Object](45839d7d-7e9b-8fe5-81f8-ee13534d3664.md)
#### Other resources


 [XmlDataBinding Object Members](ed381777-636d-df54-d2e3-9a63bebc0c6b.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/7cde1f38-aeb5-049a-361f-df1feb6dbc35.md) using GitHub.

