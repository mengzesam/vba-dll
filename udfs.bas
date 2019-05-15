Attribute VB_Name = "udfs"


Function mzsplot_Trend(figname, x, ys, tags, ysmin, ysmax, yscolor, ysdim)
        If TypeOf Application.Caller Is Range Then On Error GoTo failed
           mzsplot_Trend = Py.CallUDF("mzsPlot", "mzsplot_Trend", Array(figname, x, ys, tags, ysmin, ysmax, yscolor, ysdim), ThisWorkbook, Application.Caller)
        Exit Function
failed:
            mzsplot_Trend = Err.Description

End Function

