Public Class Control
    '---イベントプロシージャ
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '第１引数はObject型、第２引数はEventArgs型
        Button1.BackColor = Color.Pink
    End Sub

    'DataGridViewにデータソースを入れる
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dt As DataTable
        Dim dr As DataRow

        'データテーブル作成
        dt = New DataTable()

        'カラムの設定
        dt.Columns.Add(New DataColumn("ID番号", GetType(String)))
        dt.Columns.Add(New DataColumn("名前", GetType(String)))

        'データ行を追加していく
        dr = dt.NewRow()
        dr(0) = "001"
        dr(1) = "山田"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr(0) = "002"
        dr(1) = "田中"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr(0) = "003"
        dr(1) = "佐藤"
        dt.Rows.Add(dr)

        'DataSourceにDataTableを設定
        DataGridView1.DataSource = dt
    End Sub

    '入力キーの取得
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Debug.WriteLine(e.KeyChar)
    End Sub

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

    '---コントロールの取得
    Private Sub xxx()
        'Me.Controlsですべてのcontrolが格納される
        Debug.WriteLine(Me.Controls)
        'コントロールの数
        Debug.WriteLine(Me.Controls.Count)
    End Sub

    '---右クリックの判定
    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        If e.Button = MouseButtons.Right Then
            Debug.WriteLine("右クリック")
        End If
    End Sub
End Class