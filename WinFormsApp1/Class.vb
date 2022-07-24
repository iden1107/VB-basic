'クラス

'変数（フィールド）
'メソッド
'上二つを合わせてメンバという
Class Sample
    Public aaa As String = "変数" 'Dimは省略可
    Public Sub bbb()
        Console.WriteLine("メソッド")
    End Sub
End Class

'インスタンス化　sampleクラスからxxxを作り出す
Module Program
    Sub Main()
        Dim xxx = New Sample()
        xxx.bbb()　'メソッドの呼び出し
    End Sub
End Module



'---コンストラクタ
'インスタンス作成時に呼ばれるメソッド。初期化するときに使われる
Class Sample
    Public aaa As String
    'コンストラクタでaaaの値をセット
    Sub New()
        'Me.aaa = "変数" 　　この形式でもよい
        aaa = "変数"
    End Sub

    Public Sub bbb()
        Console.WriteLine("メソッド")
    End Sub
End Class
'インスタンス化
Module Program
    Sub Main()
        Dim xxx = New Sample()
        Console.WriteLine(xxx.aaa) 'aaaの"変数"が出力
    End Sub
End Module


'インスタンスの時に値をセットする
Class Sample
    Public aaa As String
    Sub New(ccc As String)　'②cccには"hoge"が格納
        aaa = ccc　'③フィールドのaaaにcccが代入
    End Sub

    Public Sub bbb()
        Console.WriteLine("メソッド")
    End Sub
End Class
'インスタンス化
Module Program
    Sub Main()
        Dim xxx = New Sample("hoge")　'①値をセット
        Console.WriteLine(xxx.aaa) '①から③を経て"hoge"が出力
    End Sub
End Module


'---オーバーロード
'パラメーターの型や数が異なれば同じメソッド名で複数定義できる
Class Sample
    Public aaa As String
    '①aaaを初期化する２つのコンストラクタがある
    '②それぞれ型が異なって定義されている
    Sub New(ccc As String)
        aaa = "文字列です"
    End Sub

    Sub New(ccc As Integer)
        aaa = "数値です"
    End Sub
End Class
Module Program
    Sub Main())
        '③インスタンス生成時は"hoge"と文字列なのでこの場合は上のコードが適用される
        'コンパイラが適用するものを分けている
        Dim xxx = New Sample("hoge")
        Console.WriteLine(xxx.aaa)
    End Sub
End Module


'オーバーロード（メソッド）
Class Sample
    'bbb()と同名のメソッドが３つあるがそれぞれ引数が異なっている
    Public Sub bbb(a As Integer)
        Console.WriteLine(a)
    End Sub

    Public Sub bbb(a As Integer, b As Integer)
        Console.WriteLine(a + b)
    End Sub
    Public Sub bbb(a As Integer, c As String)
        Console.WriteLine(a & c)
    End Sub
End Class
Module Program
    Sub Main()
        Dim xxx = New Sample()
        xxx.bbb(10, 20)　'数値が２つ渡されているので真ん中が適用
        xxx.bbb(1923, "年") '数値と文字が渡されているので下が適用 
        '同じメソッド名を使いまわせるメリットがある
    End Sub
End Module



'---継承
Public Class Sample
    Public Sub bbb()
        Console.WriteLine("メソッド")
    End Sub
End Class


Public Class SampleChild 'Sampleクラスから継承
    Inherits Sample
End Class



'---オーバーライド
Public Class Sample
    Public Overridable Sub bbb()
        Console.WriteLine("メソッド")
    End Sub
End Class

Public Class SampleChild
    Inherits Sample
    Public Overrides Sub bbb()
        Console.WriteLine("上書き")
    End Sub
End Class


'---アクセス修飾子
'Public クラス外部からも可能
'Firend　同一プログラム、ソリューション内から可能
'Protected　クラス内部、サブクラスから可能
'Private　クラス内部のみ可能


'---プロパティ
'クラス内の変数（フィールド）はprivateで宣言することが推奨される
'そのため変数にアクセスするにはメソッドを経由する方法がある
'それを便利にするのがプロパティ
Class Sample
    Private aaa As String = "文字"
End Class
Module Program
    Sub Main()
        Dim xxx = New Sample()
        Console.WriteLine(xxx.aaa)　'このままでは当然アクセスできな
    End Sub
End Module

'そこでアクセサ(get,set)を使い値を操作できるようにする
Class Sample
    Private _number As Integer = 50
    Public Property Number() As Integer
        Get
            Return _number

        End Get
        Set(value As Integer)
            _number = value
        End Set
    End Property
End Class
Module Program
    Sub Main()
        Dim xxx = New Sample()
        Console.WriteLine(xxx.Number) 'getの50が出力

        xxx.Number = 100
        Console.WriteLine(xxx.Number) '100がsetされたので100が出力
    End Sub
End Module

'上記のようにgetとsetをそのまま返すだけではpubicでフィールドを設定することとさほど変わらない
'setにロジックを加えることで値の設定に条件をつけることなどができる
Class Sample
    Private _number As Integer = 50
    Public Property Number() As Integer
        Get
            Return _number

        End Get
        Set(value As Integer) '0でない場合のみ任意の値を設定できる
            If value = Not 0 Then
                _number = value
            End If
        End Set
    End Property
End Class

'Getしかない時はReadOnlyを付ける
Class Sample
    Private _number As Integer = 0
    Public ReadOnly Property Number() As Integer
        Get
            Return _number
        End Get
    End Property
End Class

'Setしかない時はWriteOnlyを付ける
Class Sample
    Private _number As Integer = 0
    Public WriteOnly Property Number() As Integer
        Set(ByVal value As Integer)
            _number = value
        End Set
    End Property
End Class

'自動実装プロパティ　getとsetを省略できる記述方法
'ロジックが必要な場合はgetとsetを記述しなければならない
'自動実装プロパティと ReadOnly、WriteOnly キーワードの併用もできない
Class Sample
    Public Property Number() As Integer
    'Public Property Number() As Integer = 100 初期値の設定もできる

End Class
Module Program
    Sub Main()
        Dim xxx = New Sample()
        Console.WriteLine(xxx.Number)
        xxx.Number = 200
        Console.WriteLine(xxx.Number)
    End Sub
End Module



'---既定のプロパティ (C#ではインデクサ)
'クラスのインスタンスを配列のようにインデックスで呼び出せるようにしたもの
Class Sample
    Private arr() As String
    Sub New(ByVal size As Integer)
        arr = New String(size) {}
    End Sub
    Default Property Item(ByVal index As Integer) As String
        Get
            Return arr(index)
        End Get
        Set(ByVal value As String)
            arr(index) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim xxx = New Sample(3)
        xxx(1) = "hoge"
        xxx(2) = "foo"
        xxx(3) = "hogehoge"

        Console.WriteLine(xxx(2))
    End Sub
End Module


'--- moduleクラス(C#:staticクラス)
' shared修飾子(C#:static修飾子)
'モジュールは、メンバが暗黙的に Shared である型
'オブジェクト指向ではないため使用は推奨されない
'Shared修飾子を付与した変数は、インスタンスを作成しなくても存在


Class Sample
    Public Shared aaa As String = "変数"
    Public Shared Sub bbb()
        Debug.WriteLine("メソッド")
    End Sub
End Class


Sub Main()
    Console.WriteLine(Sample.aaa)
    'インスタンスを生成せずに直接使える
    bbb() 'メソッドも同様に直接呼びだし
End Sub


'---抽象クラス
'複数人で開発を行う場合に実装レベルのルールを作れる意図
'継承してもらうことを前提にしているクラスで、それ自身をインスタンス化することはできない
'継承された派生クラスでは抽象メソッドを必ずオーバーライドしなければならない制約がある
Public MustInherit Class Sample '①抽象クラス
    Public MustOverride Sub Method1()
End Class

Public Class Class　'②継承した具象クラス
    Inherits Sample
    Public Overrides Sub Method1()
        Debug.WriteLine("hoge")
    End Sub
End Class


Private Sub main()
    '③具象クラスからインスタンス化
    Dim xxx = New ClassSample()
    xxx.Method1()
End Sub
