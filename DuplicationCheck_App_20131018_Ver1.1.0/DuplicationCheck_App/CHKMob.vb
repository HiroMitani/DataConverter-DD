Imports System.Text.RegularExpressions

Module CHKMOB
    '***********************************************
    ' 数値入力チェックを行う。
    ' <引数>
    ' ChkNum : チェックをするデータを格納
    ' ChkType : チェックするデータ型を指定（INTEGER:整数数値型、DOUBLE:浮動小数点型）
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' ZeroChk : 入力したテキストボックスで0の値を許可するか（TRUE:許可、FALSE:却下）
    ' MinusChk : 入力したテキストボックスでマイナスの値を許可するか(TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function NumChkVal(ByVal ChkNum As String, _
                              ByVal ChkType As String, _
                              ByVal NullChk As Boolean, _
                              ByVal ZeroChk As Boolean, _
                              ByVal MinusChk As Boolean, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        Dim CheckInteger As Integer
        Dim CheckDouble As Double

        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkNum = "" Then
                ErrorMessage = "数量を入力してください。"
                Return False
                Exit Function
            End If
        End If

        '渡されてきた値の入力チェックタイプを判別。
        If ChkType = "INTEGER" Then
            '渡されてきた値はINT型として正しいかチェック。
            If Int32.TryParse(ChkNum, CheckInteger) Then
            Else
                ErrorMessage = "半角整数の値しか入力できません。"
                Return False
                Exit Function
            End If

            '渡されてきた値は0がOKか。
            If ZeroChk = True Then
            Else
                If ChkNum = 0 Then
                    ErrorMessage = "1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '渡されてきた値はマイナス値がOKか。
            If MinusChk = True Then
            Else
                If ChkNum < 0 Then
                    ErrorMessage = "1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '小数点でもOKの場合のチェックはDOUBLE型へ
        ElseIf ChkType = "DOUBLE" Then
            '渡されてきた値はINT型として正しいかチェック。
            If Double.TryParse(ChkNum, CheckDouble) Then
            Else
                ErrorMessage = "半角数値の値しか入力できません。"
                Return False
                Exit Function
            End If

            '渡されてきた値は0がOKか。
            If ZeroChk = True Then
            Else
                If ChkNum = 0 Then
                    ErrorMessage = "数値は1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If

            '渡されてきた値はマイナス値がOKか。
            If MinusChk = True Then
            Else
                If ChkNum < 0 Then
                    ErrorMessage = "数値は1以上の半角整数の値を入力してください。"
                    Return False
                    Exit Function
                End If
            End If
        End If
        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 日付入力チェックを行う。
    ' <引数>
    ' ChkDate : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' ChkDateAfter : 変換後のデータを格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function DateChkVal(ByVal ChkDate As String, _
                              ByVal NullChk As Boolean, _
                              ByRef ChkDateAfter As String, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean

        Dim CheckDate As Date

        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkDate = "" Then
                ErrorMessage = "日付をYYYY/MM/DD形式で入力してください。"
                Return False
                Exit Function
            End If
        End If

        '渡されてきた値はDATE型として正しいかチェック。
        If DateTime.TryParse(ChkDate, CheckDate) Then
            ChkDateAfter = CheckDate.ToString("yyyy/MM/dd")
        Else
            ErrorMessage = "入力された値は日付として正しくありません。"
            Return False
            Exit Function
        End If

        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 文字入力チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' SingleQuotationChk : シングルコーテーションを許可するか（TRUE:許可、FALSE:却下）　却下の場合は''に置き換え
    ' MaxSize : 最大文字数
    ' <戻り値>
    ' ChkStringAfter : 処理後の文字列を格納
    ' Result : True（成功） , False(失敗）
    ' ErrorMessage : エラーメッセージ
    '***********************************************
    Public Function StringChkVal(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByVal SingleQuotationChk As Boolean, _
                              ByVal MaxSize As Integer, _
                              ByRef ChkStringAfter As String, _
                              ByRef Result As String, _
                              ByRef ErrorMessage As String) As Boolean


        Dim i As Integer

        ChkStringAfter = ChkString
        '渡されてきた値がNullがOKか
        If NullChk = True Then
        Else
            If ChkString = "" Then
                ErrorMessage = "文字が入力されていません。"
                Return False
                Exit Function
            End If
        End If

        If ChkString.Length > MaxSize Then
            ErrorMessage = "20文字以上の文字列が入っています。"
            Return False
        End If
        '渡されてきた値はシングルコーテーションを許可するか。許可しない場合は変換を行う。
        If SingleQuotationChk = True Then
        Else
            i = InStr(ChkString, "'")
            '文字列の中にシングルコーテーションがあった場合はシングルコーテーションを２つ''に変換する。
            If i > 0 Then
                ChkStringAfter = ChkString.Replace("'", "''")
            End If
        End If

        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' 郵便番号チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : 入力したテキストボックスでNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' ReturnString : 変換後のChkStringを格納し戻す。
    '***********************************************
    Public Function PostDataCheck(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByRef ReturnString As String) As Boolean

        'NULLがOKか。
        If NullChk = True And ChkString = "" Then
            ReturnString = ChkString
            Return True
        Else
            If ChkString = "" Then
                Return False
            End If
        End If

        '全角から半角へ変換する。
        ChkString = StrConv(ChkString, VbStrConv.Narrow)

        '文字列が8文字以内かチェックをする。
        If Trim(ChkString.Length) > 8 Then
            ' エラー
            Return False
        End If

        'xxx-xxxx形式か
        If Regex.IsMatch(ChkString, "^[0-9]{3}[\-][0-9]{4}$") Then
            ReturnString = ChkString
            Return True
            'ハイフン無しか。（＝半角数値のみか）
        ElseIf (Regex.IsMatch(ChkString, "^[0-9]+$")) Then
            'もし7桁の場合（9999999）は、999-9999の形式に変換。
            If Trim(ChkString.Length) = 7 Then
                '前後スペースを省い格納する。
                ReturnString = Trim(ChkString)
                'ハイフンをつけて格納する。
                ReturnString = ChkString.Substring(0, 3) & "-" & ChkString.Substring(3, 4)
                Return True
            Else
                '7桁じゃなかったらFalseを返す。
                Return False
            End If
        Else
            '上記2パターン以外ならFalseを返す。
            Return False
        End If

        Return True

    End Function

    '***********************************************
    ' 電話番号チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' NullChk : ChkStringのNullを許可するか（TRUE:許可、FALSE:却下）
    ' <戻り値>
    ' ReturnString : 変換後のChkStringを格納し戻す。
    '***********************************************
    Public Function TelDataCheck(ByVal ChkString As String, _
                              ByVal NullChk As Boolean, _
                              ByRef ReturnString As String) As Boolean

        'NULLがOKか。
        If NullChk = True And ChkString = "" Then
            Return True
        Else
            If ChkString = "" Then
                Return False
            End If
        End If

        '全角から半角に変換する。
        ChkString = StrConv(ChkString, VbStrConv.Narrow)
        'ハイフンを取り除く
        ChkString = ChkString.Replace("-", "")

        '文字列が20文字以内かチェックをする。
        If Trim(ChkString.Length) > 20 Then
            ' エラー
            Return False
        End If

        '半角数値のみかチェック
        If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^[0-9]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            'もし桁数が9か10桁で、先頭が0でない場合は0を先頭に追加。

            If (Trim(ChkString.Length) = 9 Or Trim(ChkString.Length) = 10) And Trim(ChkString.Substring(0, 1)) <> 0 Then
                '前後スペースを省き、前に0を追加し格納する。
                ReturnString = "0" & Trim(ChkString)
                Return True
            Else
                ReturnString = Trim(ChkString)
                Return True
            End If
        Else
            Return False
        End If

        Return True
    End Function

    '***********************************************
    ' 会員番号チェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' <戻り値>
    ' ReturnString : 変換後のChkStringを格納し戻す。
    '***********************************************
    Public Function MemberNoDataCheck(ByVal ChkString As String, _
                              ByRef ReturnString As String) As Boolean

        '未入力ならFalseを返す。
        If ChkString = "" Then
            Return False
        End If

        '全角から半角へ変換する。
        ChkString = StrConv(ChkString, VbStrConv.Narrow)

        '文字列が20文字以内かチェックをする。
        If Trim(ChkString.Length) > 20 Then
            ' エラー
            Return False
        End If

        '半角英数字のみかチェック
        If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^[0-9a-zA-Z]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
            ReturnString = ChkString
            Return True
        Else
            Return False

        End If

        ReturnString = ChkString
        Return True
    End Function

    '***********************************************
    ' メールアドレスチェックを行う。
    ' <引数>
    ' ChkString : チェックをするデータを格納
    ' <戻り値>
    ' ReturnString : 変換後のChkStringを格納し戻す。
    '***********************************************
    Public Function MailAddressNoDataCheck(ByVal ChkString As String, _
                              ByRef ReturnString As String) As Boolean

        '未入力ならTrueを返す。
        If ChkString = "" Then
            Return True
        End If

        '文字列が200文字以内かチェックをする。
        If Trim(ChkString.Length) > 200 Then
            ' エラー
            Return False
        End If

        'MSDNを参考にし、メールアドレスが有効な物かのチェックを追加。
        'http://msdn.microsoft.com/ja-jp/library/01escwtf.aspx#Y200
        'If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$") Then
        '上記RFC準拠から()<>[]:;,の記号を許可。また連続したピリオド、＠の前のピリオドを許可。
        '@がないもの、@が2つあるものはエラー。
        '2013/04/30 　@の前の-（ハイフン）を許可。アルファベットの大文字を許可。
        'If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|\(\)\<\>\[\]\:\;\,\.~\w])*)(?<=[0-9a-zA-Z\.-])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]*\.)+[0-9a-zA-Z]{2,17}))$") Then
        '2013/09/10 @の前の_（アンダーバー）を許可。
        'If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|\(\)\<\>\[\]\:\;\,\.~\w])*)(?<=[0-9a-zA-Z\._-])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]*\.)+[0-9a-zA-Z]{2,17}))$") Then
        If System.Text.RegularExpressions.Regex.IsMatch(ChkString, "^(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|\(\)\<\>\[\]\:\;\,\.~\w])*)(?<=[0-9a-zA-Z\._-])@))" + "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]*\.)+[a-z0-9]{2,17}))$") Then


            ReturnString = ChkString
            Return True
        Else
            Return False
        End If

    End Function
End Module
