Attribute VB_Name = "mod1"
Sub e()

dim a as integer
a=1
a=foo(a)
MsgBox (a)

End Sub


' ���Ժ���˵�� with ����

function bar(j as integer)

bar=j+2

end function


' ���Ժ���˵�� with ����
function foo(i as integer)

foo=i+1

End function


