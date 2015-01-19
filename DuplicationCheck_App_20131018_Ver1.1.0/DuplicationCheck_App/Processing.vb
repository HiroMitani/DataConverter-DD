Public Class Processing

    Private Sub Processing_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim h, w As Integer
        'ディスプレイの高さ
        h = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height
        'ディスプレイの幅
        w = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point((w / 2) - 150, (h / 2) - 75)
        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle
    End Sub
End Class