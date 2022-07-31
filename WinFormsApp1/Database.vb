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



'---取得
Imports System.Data.SqlClient

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn = New SqlConnection("Data Source=DESKTOP-RLPQ3N9\SQLEXPRESS;Initial Catalog=sample;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
        Dim SQLCm As SqlCommand = cn.CreateCommand

        SQLCm.CommandText = "SELECT * FROM Table_1 WHERE Age = 34"
        Dim Value As String

        cn.Open()
        Value = SQLCm.ExecuteScalar
        cn.Close()

        MsgBox(Value)
    End Sub
End Class


'---追加
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn = New SqlConnection("Data Source=DESKTOP-RLPQ3N9\SQLEXPRESS;Initial Catalog=sample;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
        Dim SQLCm As SqlCommand = cn.CreateCommand

        SQLCm.CommandText = "INSERT INTO Table_1 VALUES ('山口',55) "
        Dim Value As String

        cn.Open()
        Value = SQLCm.ExecuteNonQuery()
        cn.Close()

        MsgBox(Value)
    End Sub
End Class


'---更新
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn = New SqlConnection("Data Source=DESKTOP-RLPQ3N9\SQLEXPRESS;Initial Catalog=sample;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
        Dim SQLCm As SqlCommand = cn.CreateCommand

        SQLCm.CommandText = "UPDATE Table_1 SET Age = 100 WHERE Name = '山田'"
        Dim Value As String

        cn.Open()
        Value = SQLCm.ExecuteNonQuery()
        cn.Close()

    End Sub
End Class



'---削除
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn = New SqlConnection("Data Source=DESKTOP-RLPQ3N9\SQLEXPRESS;Initial Catalog=sample;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False")
        Dim SQLCm As SqlCommand = cn.CreateCommand

        SQLCm.CommandText = "DELETE FROM Table_1　WHERE Name = '山口'"
        Dim Value As String

        cn.Open()
        Value = SQLCm.ExecuteNonQuery()
        cn.Close()

    End Sub
End Class