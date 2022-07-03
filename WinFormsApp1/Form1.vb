Public Class Form1

    Private Sub fname()

        'メソッド
        TextBox1.Focus()
        TextBox1.SelectAll()

        'プロパティ
        TextBox1.Text = "Hello World"
        TextBox1.BackColor = Color.Red

        '文字の右寄せ
        TextBox1.TextAlign = HorizontalAlignment.Right

        '文字列の結合
        Dim aaa, bbb As String
        aaa = "hello"
        bbb = "world"
        MsgBox(aaa & bbb) '+でもokだが&推奨
        MsgBox($"{aaa}!!!{bbb}!!!!")

        '数値の扱い
        Dim x As Decimal = 815.6D '後ろのDはデジマルであることを明示する。型を明確にするほうがよい

        '日付
        Dim day As Date = #6/27/2020#
        MsgBox(day) '書式指定なし。Windowsの設定で書式化されます。
        MsgBox(day.ToString("yyyy年MM月dd日")) '2020年06月27日
        MsgBox(day.ToString("yy-MM-dd")) '20-06-27
        day = Now '現在日付時刻
        MsgBox(Now.ToString("HH:mm:ss")) '時刻

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim day As Date = Now
        MsgBox(Now.ToString("HH:mm:ss"))
    End Sub
End Class


