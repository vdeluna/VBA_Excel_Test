
# ODBCConnection.SourceDataFile Property (Excel)

Returns or sets a  **String** indicating the source data file for an ODBC connection. Read/write.


## Version Information

Version Added: Excel 2007 


## Syntax

 _expression_. **SourceDataFile**

 _expression_A variable that represents an  **ODBCConnection** object.


## Remarks

For file-based data sources (for example, Access) the  **SourceDataFile** property contains a fully qualified path to the source data file. It is null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to null if the ** [Connection](2fcd1043-b088-cfde-9853-4a20da20be26.md)** property is changed programmatically.


## See also


#### Concepts


 [ODBCConnection Object](b880ebec-15a4-5a3d-ef02-db73106db9c9.md)
#### Other resources


 [ODBCConnection Object Members](d13b91f3-a89f-7dd7-7a98-f1d952f3b047.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/f32c0eeb-e8f5-1a9f-13fd-ead4ad96381f.md) using GitHub.

