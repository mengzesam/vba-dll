Attribute VB_Name = "xlwings_udfs"


Function mzsplot_Trend(figname, x, ys, tags, ysmin, ysmax, yscolor, ysPhiFlag, drawStyle)
        If TypeOf Application.Caller Is Range Then On Error GoTo failed
           mzsplot_Trend = Py.CallUDF("mzsPlot", "mzsplot_Trend", Array(figname, x, ys, tags, ysmin, ysmax, yscolor, ysPhiFlag, drawStyle) _
               , ThisWorkbook, Application.Caller)
        Exit Function
failed:
            mzsplot_Trend = Err.Description

End Function

