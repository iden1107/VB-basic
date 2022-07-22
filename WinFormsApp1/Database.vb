Public Class Database
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '接続するクラス
        'OleDbConnection どのデータベースを使用するにしても、コードの記述方法は同じため、データベース切り替え時のメンテナンスが楽。
        'SqlConnection	 SQL Server専用のクラスのため、パフォーマンスが向上する。

        'SqlConnectionクラスの新しいインスタンスを初期化
        '接続文字列を引数に入れる
        Dim cnn = New System.Data.SqlClient.SqlConnection("Data Source=DESKTOP-RLPQ3N9\SQLEXPRESS;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")

        'データベース接続を開く
        cnn.Open()
        Debug.WriteLine("ok")

        'データベースへの接続を閉じる
        cnn.Close()
    End Sub
End Class