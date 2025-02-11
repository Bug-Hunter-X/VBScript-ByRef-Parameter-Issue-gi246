Function f(ByRef a)
  a = a + 1
end function

x = 10
f x
MsgBox x

'Solution: Use Set to explicitly assign the modified value
Function g(ByRef a)
  Set a = CreateObject("Scripting.Dictionary")
  a.Add "value", a + 1
end Function

Dim x
x = 10
g x
MsgBox x 