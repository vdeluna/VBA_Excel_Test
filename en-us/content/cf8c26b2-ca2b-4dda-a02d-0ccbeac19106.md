
# TempErr/TempErr12

 **Last modified:** March 09, 2015

 _**Applies to:** Excel 2013 | Office 2013 | Visual Studio_

Framework library function that creates a temporary  **XLOPER**/ **XLOPER12** containing a Microsoft Excel worksheet error.


```C#

LPXLOPER TempErr(WORD err);
LPXLOPER12 TempErr12(BOOL err);
```


## Parameters

err

The desired error code, or its literal numeric equivalent, as shown in the following table.



|**Error**|**Error code defined in XLCALL.H**|**Decimal equivalent**|
|:-----|:-----|:-----|
|#NULL| **xlerrNull**|0|
|#DIV/0!| **xlerrDiv0**|7|
|#VALUE!| **xlerrValue**|15|
|#REF!| **xlerrRef**|23|
|#NAME?| **xlerrName**|29|
|#NUM!| **xlerrNum**|36|
|#N/A| **xlerrNA**|42|

## Return value

Returns an  **xltypeBool** containing the error code passed in.


## Example

This example uses the  **TempErr12** function to return a #VALUE! error to Excel.


**Note**  The Framework library function  **TempErr12** allocates memory from an internal buffer, which is normally freed when the Framework function **Excel12f** is called. If this example function is called repeatedly without **Excel12f** being called, a memory leak occurs.

 `\SAMPLES\EXAMPLE\EXAMPLE.C`




```C#
LPXLOPER WINAPI TempErrExample(void)
{
    return TempErr12(xlerrValue);
}
```


## See also


#### Concepts


 [Functions in the Framework Library](7d9a13fd-9a4c-423e-bb08-4a5be57c7905.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/cf8c26b2-ca2b-4dda-a02d-0ccbeac19106.md) using GitHub.

