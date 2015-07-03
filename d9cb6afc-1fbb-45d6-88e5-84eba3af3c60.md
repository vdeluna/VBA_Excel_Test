
# TempMissing/TempMissing12

 **Last modified:** July 01, 2011

 _**Applies to:** Excel 2013 | Office 2013 | Visual Studio_

Framework library function that creates a temporary  **XLOPER**/ **XLOPER12** of type **xltypeMissing**.


```C#

LPXLOPER TempMissing(void);
LPXLOPER12 TempMissing12(void);
```


## Parameters

This function takes no parameters.


## Return value

Returns a pointer to an  **xltypeMissing** **XLOPER**/ **XLOPER12**.


## Example

This example uses  **TempMissing12** to provide three missing arguments to **xlcWorkspace** followed by a **Boolean** **FALSE** to suppress the display of worksheet scroll bars. The first three arguments correspond to other workspace settings which are unaffected.

 `\SAMPLES\EXAMPLE\EXAMPLE.C`




```C#
short WINAPI TempMissingExample(void)
{
   XLOPER12 xBool;

   xBool.xltype = xltypeBool;
   xBool.val.xbool = 0;
   Excel12f(xlcWorkspace, 0, 4, TempMissing12(), TempMissing12(),
      TempMissing12(), (LPXLOPER12)&amp;xBool);
   return 1;
}
```


## See also


#### Concepts


 [Functions in the Framework Library](7d9a13fd-9a4c-423e-bb08-4a5be57c7905.md)
****   **Contribute to this article**Want to edit or suggest changes to this content? You can edit and submit changes to  [this article](https://github.com/jhershey00/VBA_Excel_Test/OpenXMLCon/articles/d9cb6afc-1fbb-45d6-88e5-84eba3af3c60.md) using GitHub.

