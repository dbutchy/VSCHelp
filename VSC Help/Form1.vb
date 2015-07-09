Public Class frmMain

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox_VSC(ex.Message, ERROR_TITLE_NORMAL)
        End Try
    End Sub
End Class
