Option Explicit

Sub 加载所有自定义函数说明()
    Application.MacroOptions _
        "Hrz_UNIQUE", _
        "将一个区域去重，并返回第index个结果，若index超出范围则报错NA", _
        , , , , 5, "Hrz去重函数", , , Array("要去重的范围", "要输出第几个结果")
    Application.MacroOptions _
        "Hrz_SORTBY", _
        "将一组数据排序，并返回相应的另一列的排序结果的第index个名次" & Chr(10) & "例如按照“年龄”为“姓名”排序" & Chr(10) & "名次不会重合，若index超出范围则报错NA", _
        , , , , 5, "Hrz排序函数", , , Array("排序依据的值区域", "输出结果所在的区域", "要输出第几个结果", "False,降序;True,升序")
    Application.MacroOptions _
        "Hrz_SORTBY1", _
        "将一组数据排序，并返回相应的另一列的排序结果的第index个名次，该函数忽略最后连续的0值" & Chr(10) & "例如按照“年龄”为“姓名”排序" & Chr(10) & "名次不会重合，若index超出范围则报错NA", _
        , , , , 5, "Hrz排序函数", , , Array("排序依据的值区域", "输出结果所在的区域", "要输出第几个结果", "False,降序;True,升序")
End Sub


