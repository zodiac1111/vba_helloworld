Attribute VB_Name = "mod1"
Sub e()

dim a as integer
a=1
a=foo(a)
MsgBox (a)

End Sub


' 测试函数说明 with 中文

function bar(j as integer)

bar=j+2

end function


' 测试函数说明 with 中文
function foo(i as integer)

foo=i+1

End function


