Public Class Form2

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Process.Start("http://tangorin.com/general/" & Form1.clickedName2)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Process.Start("http://dictionary.goo.ne.jp/srch/all/" & Form1.clickedName2 & "/m0u/")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Process.Start("http://jisho.org/search/" & Form1.clickedName2)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Process.Start("http://www.romajidesu.com/dictionary/meaning-of-" & Form1.clickedName2 & ".html")
    End Sub
End Class