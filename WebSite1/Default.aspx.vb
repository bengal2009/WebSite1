
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Calendar1_SelectionChanged(sender As Object, e As EventArgs) Handles Calendar1.SelectionChanged
        MsgBox(Me.Calendar1.SelectedDate)

    End Sub

    Private Sub form1_Load(sender As Object, e As EventArgs) Handles form1.Load


    End Sub
End Class
