
# Multithreading and Memory Management

 **Last modified:** November 07, 2011

 _**Applies to:** Excel 2013 | Office 2013 | Visual Studio_

Proper handling of memory is vital to creating reliable XLL add-ins for Microsoft Excel. Failure to allocate appropriate memory buffers and free them when they are no longer needed reduces performance, creates resource contention, and destabilizes Excel.

Beginning with Microsoft Office Excel 2007, you can configure Excel to use up to 1,024 concurrent threads when recalculating. In some cases, especially when multiple processors are available or with user-defined functions running on clustered servers, multithreading can improve performance.
The following topics describe how to manage memory and threads in XLLs:

-  [Memory Management in Excel](3bf5195b-6235-43cf-8795-0c7b0a63a095.md)
    
-  [Multithreading and Memory Contention in Excel](86e1e842-f433-4ea9-8b13-ad2515fc50d8.md)
    
-  [Multithreaded Recalculation in Excel](c6c831f1-4be1-4dcc-a0fa-c26052ec53c9.md)
    

## See also


#### Concepts


 [Developing Excel 2013 XLLs](dd27ae4d-ef97-47db-885c-ddd955816900.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/6f7e052a-4270-4b83-b1ed-feabf6dbeaa2.md) using GitHub.

