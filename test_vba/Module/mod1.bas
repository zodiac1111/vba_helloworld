Attribute VB_Name = "mod1"
Sub e()

Dim a As Integer
a = 1
a = foo(a)
MsgBox (a)

End Sub


' 测试函数说明 with 中文

Function bar(j As Integer)

bar = j + 2

End Function


' 测试函数说明 with 中文
Function foo(i As Integer)

foo = i + 1

Xdebug.printx foo

End Function




