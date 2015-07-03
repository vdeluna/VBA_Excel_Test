
# FullSeriesCollection.Application Property (Excel)

Returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.


## Version information

Version Added: Excel 2013 


## Syntax

 _expression_. **Application**

 _expression_A variable that represents a  [FullSeriesCollection Object (Excel)](5d7b7e7c-0a74-307b-84f9-56143ceba464.md) object.


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


 [FullSeriesCollection Object Members](18060b3a-f25c-fa99-d3f3-dd59f7928465.md)
 [FullSeriesCollection Object](5d7b7e7c-0a74-307b-84f9-56143ceba464.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/52dfb5aa-c6fb-201c-c1ed-880aff1efb45.md) using GitHub.

