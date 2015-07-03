
# OLEDBError.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**.


## Syntax

 _expression_. **Creator**

 _expression_A variable that represents an  **OLEDBError** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## See also


#### Concepts


 [OLEDBError Object](6bcbf721-f2c8-f784-361b-e1a298bb2ecb.md)
#### Other resources


 [OLEDBError Object Members](52181252-dd6f-b267-fa21-4ad8175b7346.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/dab01efa-cefb-1dee-847e-56688ef88c14.md) using GitHub.

