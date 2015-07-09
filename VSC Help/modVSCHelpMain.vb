Module modVSCHelpMain

    Private Const mstrModuleName As String = " modVSCHelpMain "

    Public Const ERROR_TITLE_NORMAL As String = "NOTE!"

    Public Sub MsgBox_VSC(ByVal sErrMsg As String, Optional sErrTitle As String = ERROR_TITLE_NORMAL)
        Try
            MsgBox(sErrMsg, MsgBoxStyle.OkOnly, sErrTitle)
        Catch ex As Exception
            Try
                MsgBox(ex.Message & " -- while trying to post the following ERROR MESSAGE: '" & sErrMsg & "'", MsgBoxStyle.OkOnly, "TAKE SPECIAL NOTICE!")
            Catch ex1 As Exception
                MsgBox(ex1.Message, MsgBoxStyle.OkOnly, "TAKE SPECIAL NOTICE!")
            End Try
        End Try
    End Sub
End Module
