Public Class Control
    '---イベントプロシージャ
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '第１引数はObject型、第２引数はEventArgs型
        Button1.BackColor = Color.Pink
    End Sub

    '入力キーの取得

    '---タイマー
    Dim n As Integer = 0　'変数を宣言

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Timer1.Enabled = True 'タイマーを起動
        Timer1.Interval = 1000　'インターバルを設定
        Label2.Text = n

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        '設定した秒ごとに処理する内容
        n += 1 'nを１増やす
        Label1.Text = n.ToString() 'ラベルの表示内容を更新
    End Sub
End Class