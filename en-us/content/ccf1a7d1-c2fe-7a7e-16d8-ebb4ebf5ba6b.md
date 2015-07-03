
# XlEnableCancelKey Enumeration (Excel)

Specifies how Microsoft Office Excel 2007 handles CTRL+BREAK (or ESC or COMMAND+PERIOD) user interruptions to the running procedure.


## Version Information

Version Added: Excel 2007 



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **xlDisabled**|0|Cancel key trapping is completely disabled.|
| **xlErrorHandler**|2|The interrupt is sent to the running procedure as an error, trappable by an error handler set up with an On Error GoTo statement. The trappable error code is 18.|
| **xlInterrupt**|1|The current procedure is interrupted, and the user can debug or end the procedure.|

****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/ccf1a7d1-c2fe-7a7e-16d8-ebb4ebf5ba6b.md) using GitHub.

