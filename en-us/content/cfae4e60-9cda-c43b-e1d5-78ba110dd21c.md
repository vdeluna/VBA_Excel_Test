
# CategoryCollection.Application Property (Excel)

Returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.


## Version information

Version Added: Excel 2013 


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  [CategoryCollection Object (Excel)](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md) object.


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


## Property value

 **APPLICATION**


## See also


#### Other resources


 [CategoryCollection Object Members](39a6f85c-2219-79df-cbbc-0bcc21a517e8.md)
 [CategoryCollection Object](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/cfae4e60-9cda-c43b-e1d5-78ba110dd21c.md) using GitHub.

