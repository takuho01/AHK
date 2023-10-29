Array := Object()
Array[0] := "item"
Array[1] := 1
Array[2] := 3

MsgArrayTest1() ; 結果：{空白}, {空白}
MsgArrayTest1()
{
	Array[1] := Array[1] + 1
	MsgBox % Array[1]
	MsgBox % Array[0]
}

MsgArrayTest2() ; 結果：2, item
MsgArrayTest2()
{
	global Array ; グローバル変数を参照するための宣言
	Array[1] := Array[1] + 1
	MsgBox % Array[1]
	MsgBox % Array[0]
}

MsgVariable(Array[0]) ; 結果： item
MsgVariable(outString)
{

	MsgBox %outString%
}

MsgArray(Array) ; 結果：3
MsgArray(outString)
{
	MsgBox % outString[2] 
}