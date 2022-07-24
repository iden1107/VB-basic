'---インターフェイス
'中身の処理は記述されない
'クラスで実装(記述)することで動作する
'継承先でメソッドのオーバーライドを強制
Interface ISample
    Sub ShowMessage()
End Interface

Public Class Sample
    Implements ISample 'インターフェースを実装
    Public Sub ShowMessage() Implements ISample.ShowMessage
        Debug.WriteLine("派生クラス")
    End Sub
End Class



'---ジェネリック
'引数の型を任意のものが入る
'型を指定した例では、型が異なる場合エラーになる
'クラスをインスタンス化した際、型だけ異なる場合にその度クラスを作らなくてもよくなることがメリット

'エラーになるケース
Class Aaa
    Public bbb As Integer
End Class

Public Sub Fun1()
    Dim xxx = New Aaa()
    xxx.bbb = 100
    Debug.WriteLine(xxx.bbb) '100で作ったインスタンスではOKでも
End Sub

Public Sub Fun2()
    Dim xxx = New Aaa()
    xxx.bbb = "hoge"
    Debug.WriteLine(xxx.bbb) '文字のインスタンスはエラーになる
End Sub


'ジェネリックを使用した場合
Class Aaa(Of T)
    Public bbb As T
End Class


Public Sub Afun()
    Dim xxx As New Aaa(Of String)
    xxx.bbb = "hoge"
    Debug.WriteLine(xxx.bbb)
End Sub

Public Sub Bfun()
    Dim xxx As New Aaa(Of Integer)
    xxx.bbb = 100
    Debug.WriteLine(xxx.bbb)
End Sub



'---名前空間
'クラス名の衝突や、クラスの分類などに使える
Namespace Aaa
    Public Class Bbb
        Public Sub Ccc()
            Debug.WriteLine("Aaa名前空間のCccメソッド実行")
        End Sub
    End Class

    Public Class Ddd
        Public Sub Eee()
            Debug.WriteLine("Aaa名前空間のDddメソッド実行")
        End Sub
    End Class
End Namespace


Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xxx = New Aaa.Bbb()
        xxx.Ccc()
    End Sub
End Class

'---importによる名前空間
Imports WinFormsApp1.Aaa

Namespace Aaa
    Public Class Bbb
        Public Sub Ccc()
            Debug.WriteLine("Aaa名前空間のCccメソッド実行")
        End Sub
    End Class

    Public Class Ddd
        Public Sub Eee()
            Debug.WriteLine("Aaa名前空間のDddメソッド実行")
        End Sub
    End Class
End Namespace


Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'ここで名前空間を省略できる
        Dim xxx = New Bbb()
        xxx.Ccc()
    End Sub
End Class



'---構造体
'クラス（参照型）と似ている
'値型。継承などができない
'軽量のオブジェクトを表すのに適している
'小規模のデータであれば、処理が早い
'継承、イベントなどができない

Private Structure Users
    Public UserID As String
    Public Name As String
    Public Age As Integer
    Public Sub setUsers(ByVal setVal1 As String, ByVal setVal2 As String, ByVal setVal3 As Integer)
        UserID = setVal1
        Name = setVal2
        Age = setVal3
    End Sub
End Structure

Shared Sub Main(args As String())
    '構造体を宣言し、初期化
    Dim sUsers As Users
    sUsers.setUsers("A0001", "テストユーザ1", 24)
End Sub



'---列挙型
'複数の定数を1つにまとめる
'列名に対応した数値を割り当てることで数値として扱える

Public Enum Color
    RED
    YELLOW
    BLUE
End Enum

Module Module1
    Sub Main()
        Console.WriteLine(Color.RED) '0が出力される
        Console.WriteLine(Color.YELLOW) '1が出力される
        Console.WriteLine(Color.BLUE) '2が出力される
    End Sub
End Module

'最初の数値を１から始めたい場合
Enum Color
    RED = 1
    YELLOW
    BLUE
End Enum

Module Module1
    Sub Main()
        Console.WriteLine(Color.RED) '1が出力される
        Console.WriteLine(Color.YELLOW) '2が出力される
        Console.WriteLine(Color.BLUE) '3が出力される
    End Sub
End Module

'数字の割り当てを任意の値にする場合
Enum Color
    RED = 15
    YELLOW = -2
    BLUE = 542
End Enum

'数値ではなく名称で出力する場合
Enum Color
    RED
    YELLOW
    BLUE
End Enum

Module Program
    Sub Main(args As String())
        Console.WriteLine(Color.RED.ToString)
    End Sub
End Module

'名称をループで取り出す
Module Program
    Sub Main(args As String())
        For Each s In [Enum].GetNames(GetType(Color))
            Console.WriteLine(s)
        Next
    End Sub
End Module

'数値をループで取り出す
Module Program
    Sub Main(args As String())
        For Each i In [Enum].GetValues(GetType(Color))
            Console.WriteLine(i) '0 1 2
        Next
    End Sub
End Module



'---デリゲート
'メソッド(関数)を変数のように扱う機能
'デリゲートは”型”をつくることがポイント
Public Class Form1
    'デリゲートの宣言
    Delegate Sub sample()

    '関数を用意
    Sub Aaa()
        Debug.WriteLine("Aaaメソッド")
    End Sub

    Sub Bbb()
        Debug.WriteLine("Bbbメソッド")
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '変数を用意し、関数を割り当てる
        Dim Ccc As sample = AddressOf Aaa　'この時点では実行しないため　Bbb()としない
        '変数で実行する
        Ccc()
    End Sub
End Class

'引数ありの場合
'注　デリゲートと型と代入する関数の形式を同一にすること！
Public Class Form1
    'デリゲートの宣言
    Delegate Function sample(arg As Integer) As Integer

    '関数を用意
    Public Function Aaa(arg As Integer)
        Return arg * 10 '引数を10倍にするメソッド
    End Function

    Function Bbb(arg As Integer)
        Return arg + 10  '10増やすメソッド
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '変数を用意し、関数を割り当てる
        Dim Ccc As sample = AddressOf Aaa
        Debug.WriteLine(Ccc(2))
    End Sub
End Class

'マルチキャストデリゲート
'一つのデリゲート変数に複数のデリゲートを登録
Public Class Form1
    'デリゲートの宣言
    Delegate Sub sample()

    '関数を用意
    Sub Aaa()
        Debug.WriteLine("Aaaメソッド")
    End Sub

    Sub Bbb()
        Debug.WriteLine("Bbbメソッド")
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '変数を用意し、関数を割り当てる
        Dim Ccc As sample = AddressOf Aaa
        Dim Ddd As sample = AddressOf Bbb
        'combineを使いCccとDddをCcc()にまとめることで複数のデリゲートを登録
        Ccc = System.Delegate.Combine(Ccc, Ddd)
        Ccc()
    End Sub
End Class



'---ラムダ式
'VBには匿名メソッドを作成する機能なし
'省略した形式で記述できる

'通常の形式
Public Class Form1
    'デリゲートの宣言
    Delegate Sub Aaa()

    '関数の用意
    Sub Bbb()
        Debug.WriteLine("hoge")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '用意した関数をCccに代入
        Dim Ccc As Aaa = AddressOf Bbb
        '実行
        Ccc()
    End Sub
End Class

'ラムダ式
Public Class Form1
    'デリゲートの宣言
    Delegate Sub Aaa()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '代入する関数は用意せず、右辺に直接記述
        'なのでaddressofが不要
        Dim Ccc As Aaa = Sub() Debug.WriteLine("hoge")
        '実行
        Ccc()
    End Sub
End Class

'引数ありver
Public Class Form1
    'デリゲートの宣言
    Delegate Sub Aaa(arg As String)

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Ccc As Aaa = Sub(x) Debug.WriteLine(x)
        '実行
        Ccc("hoge")
    End Sub
End Class



'---LINQ
'データの集合から取得したい情報を引っ張ってくるための仕組み
Public Class Form1
    Dim aaa() As Integer = {60, 30, 80}

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'リストの中から50以上のものを抽出
        Dim result = From a In Aaa
                     Where a >= 50

        'ループで出力
        For Each r In result
            Debug.WriteLine(r)
        Next
    End Sub

End Class