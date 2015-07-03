
# xlfRegisterId

 **Last modified:** July 01, 2011

 _**Applies to:** Excel 2013 | Office 2013 | Visual Studio_

Can be called from a DLL that has itself been called by Microsoft Excel. If a function is already registered, it returns the existing register ID for that function without reregistering it. If a function is not yet registered, it registers it and returns the resulting register ID.


```C#

Excel12(xlfRegisterId, LPXLOPER12 pxRes, 3,     LPXLOPER12 pxModuleText, LPXLOPER12 pxProcedure, LPXLOPER12 pxTypeText);
```


## Parameters

pxModuleText ( **xltypeStr**)

The name of the DLL containing the function.

pxProcedure ( **xltypeStr** or **xltypeNum**)

If a string, the name of the function to call. If a number, the ordinal export number of the function to call. For clarity and robustness, always use the string form.

pxTypeText ( **xltypeStr**)

An optional string specifying the types of all the arguments to the function and the type of the return value of the function. For more information, see the "Remarks" section. This argument can be omitted for a stand-alone DLL (XLL) defining  **xlAutoRegister**.


## Property Value/Return Value

Returns the register ID of the function ( **xltypeNum**), which can be used in subsequent calls to  **xlfUnregister**.


## Remarks

This function is useful when you do not want to worry about maintaining a register ID, but you need one later for unregistering. It is also useful for assigning to menus, tools, and buttons when the function you want to assign is in a DLL.

Where a DLL or XLL function has been registered with a valid pxFunctionText argument having been supplied to **xlfRegister**, its register ID can also be obtained by passing the pxFunctionText to the function **xlfEvaluate**.


## See also


#### Reference


 [REGISTER](c730124c-1886-4a0f-8f06-79763025537d.md)
 [UNREGISTER](850bf65f-a151-44d6-b49f-d53ae2c83760.md)
#### Concepts


 [Essential and Useful C API XLM Functions](dc80cb3d-0d7e-4cb9-9870-3acc84eeca82.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/d34cf20c-a5cd-45fb-9dcb-d49eac2d99dd.md) using GitHub.

