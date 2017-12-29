# Optional

Represents arguments that the VBScript client doesn't need to specify.

## Receiving optional parameters in C++ from VBScript

Here'll you see that VBScript passes in optional argument as `VT_ERROR`.

```c++
bool VBScript_IsOptional(VARIANT* value)
{
  return value && value->vt == VT_ERROR && V_ERROR(value) == DISP_E_PARAMNOTFOUND;
}
```

## Sending optional parameters from C++ to VBScript

You'll see that the optional parameter is actually typed `VT_ERROR` instead of `VT_NULL` or `VT_EMPTY`.

```c++
void VBScript_SetOptional(VARIANT* value)
{
  VariantInit(value);
  value->vt = VT_ERROR;
  V_ERROR(value) = DISP_E_PARAMNOTFOUND;
}
```

## See Also

 [MSDN 3.1.4.4.3 - Handling Default Value and Optional Arguments](https://msdn.microsoft.com/en-us/library/cc237626.aspx)
