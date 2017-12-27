Use the following code to return `Nothing` to `VBScript`:

```c++
void VBScriptNothing(VARIANT* value)
{
  VariantInit(value);
  value->vt = VT_DISPATCH;
  V_DISPATCH(&value) = NULL;
}
```
