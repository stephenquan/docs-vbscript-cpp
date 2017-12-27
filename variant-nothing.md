# Nothing

Represents an object in VBScript that contains no value. Setting objects to this value can be used to release COM objects.

## VBScript

The following causes `msxml3.dll` to be loaded and the `IXMLDOMDocument` COM object to be created and used. Setting the object to `Nothing` allows the COM object to be released. Should `CoFreeUnusedLibrary` be called, `msxml3.dll` can be unloaded.

```vbscript
Set dom = CreateObject("Microsoft.XMLDOM")
dom.loadXML "<Hello><World/></Hello>"
Set dom = Nothing
```

## Receiving Nothing in C++ from VBScript

Here'll you see that VBScript passes in the `Nothing` object with as `VT_DISPATCH`.

```c++
bool VBScript_IsNothing(VARIANT* value)
{
  return value && value->vt == VT_DISPATCH && V_DISPATCH(value) == NULL;
}
```

## Returning Nothing from C++ to VBScript

You'll see that the `Nothing` object is actually typed `VT_DISPATCH` instead of `VT_NULL` or `VT_EMPTY`. This furthers the idea that `Nothing` is an object. Also, it furthers the idea that we prefer to return `VT_DISPATCH` to VBScript instead of `VT_UNKNOWN`.

```c++
void VBScript_SetNothing(VARIANT* value)
{
  VariantInit(value);
  value->vt = VT_DISPATCH;
  V_DISPATCH(&value) = NULL;
}
```
