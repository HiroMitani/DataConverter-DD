Imports System.IO
Imports System
Imports System.Text
Imports System.Data.OleDb
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.Globalization

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim MjstrPath As String
    'DBのpath
    Dim strMdbpath As String
    '取り込みファイルのディレクトリを格納
    Dim CSVDirName As String
    '取り込みファイルの拡張子を格納
    Dim FileExtension As String
    'カラム名一覧
    'Dim DBColumnList() As String = Nothing
    'カラム数を格納
    Dim DBColumnCount As Integer = 0

    Dim ErrorDataItemName() As String = Nothing
    Dim ErrorDataNo() As String = Nothing
    Dim ErrorDataContent() As String = Nothing
    'エクセルの一行の項目数をカウントする。
    Public TitleItemCount As Integer = 0
    'DD取り込みファイルフォーマットの項目名称
    Dim FormatList() As String = {"登録店舗名", "会員番号", "顧客名(姓)", "顧客名(名)", "ふりがな(姓)", _
                                  "ふりがな(名)", "郵便番号1", "住所1", "郵便番号2", "住所2", _
                                  "電話番号(自宅)", "電話番号(携帯)", "電話番号(会社)", "電話番号(その他)", "性別", _
                                  "誕生日", "初回来店日", "PCﾒｰﾙｱﾄﾞﾚｽ", "PCﾒｰﾙ受信許可", "携帯ﾒｰﾙｱﾄﾞﾚｽ", _
                                  "携帯ﾒｰﾙ受信許可", "主担当", "DM許可フラグ", "整理番号", "会員メモ", _
                                  "反響日", "媒体名", "反響電話担当者", "反響メモ", "過去の来店回数", _
                                  "ランク"}

    '***********************************************
    ' NumberからA,Bのようにエクセルの横軸に対応するアルファベットを求める。
    ' <引数>
    ' Number : 求める為の数値
    ' <戻り値>
    ' ItemName : 求めたアルファベットを格納
    ' Result : True（成功） , False(失敗）
    '***********************************************
    Public Function ItemNameChk(ByVal Number As Integer, _
                                ByRef ItemName As String, _
                                ByRef Result As String) As Boolean
        Dim Hi As Integer
        Dim Low As Integer
        'Numberが空ならFalseを返す。
        If IsDBNull(Number) Then
            Return False
        End If

        Hi = Number \ 26
        Low = (Number Mod 26) + 1
        If Hi <> 0 Then
            '@に求まった数値をプラスして格納。Hiが0ではない場合、AAなど二文字となるので、商と余りの値をそれぞれ使用する。
            ItemName = Chr(&H40 + Hi) & Chr(&H40 + Low)
        Else
            '@に求まった数値をプラスして格納。
            ItemName = Chr(&H40 + Low)
        End If

        '全てのチェックがOKならTrueを返す
        Return True
    End Function

    '***********************************************
    ' Excel作成時の参照の解放
    ' <引数>
    'なし
    ' <戻り値>
    ' Result : True（成功） , False(失敗）
    '***********************************************
    Public Function Excel_Release(ByVal range, ByVal range1, ByVal range2, ByVal oBook2, ByVal oBooks2, ByVal oSheet2, ByVal oSheets2, ByVal oExcel2) As Boolean

        '参照を解放(range,range1,range2)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(range)
        range = Nothing
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(range1)
        range1 = Nothing
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(range2)
        range2 = Nothing
        '参照を解放(oBook)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook2)
        'oBookを解放
        oBook2 = Nothing
        '参照を解放(oBooks)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBooks2)
        oBooks2 = Nothing
        '参照を解放(oSheet)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheet2)
        oSheet2 = Nothing
        '参照を解放(oSheets)
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheets2)
        oSheets2 = Nothing
        '参照を解放(oExcel)
        oExcel2.Quit()
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel2)
        oExcel2 = Nothing
        Return True
    End Function

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Formのロード
        Dim strWorkPath As String

        'アプリケーションディレクトリを取得する
        strWorkPath = _
        System.Reflection.Assembly.GetExecutingAssembly.Location
        Dim myFileInfo As New System.IO.FileInfo(strWorkPath)
        MjstrPath = myFileInfo.DirectoryName

        'マイドキュメントのパス
        TextBox2.Text = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal)

        'フォームの最大化を非表示にする。
        Me.MaximizeBox = Not Me.MaximizeBox

        'フォームの表示位置の設定
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(0, 0)

        'ウインドウのサイズ変更不可能にする。
        FormBorderStyle = FormBorderStyle.FixedSingle

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'CSVファイルの指定を行い、データの読み込み、DBの作成、データの取り込みを行う。
        '処理開始時刻を記憶
        'Dim start As DateTime = Now
        Dim Result As Boolean = True
        '１行のデータ項目数を格納
        Dim intCnt As Integer

        'データ数をカウント
        Dim DataCount As Integer = 0
        Dim Count As Integer = 0

        'エクセルの一行の項目数をカウントする。
        Dim TitleItemCount As Integer = 0

        Dim ItemCountCheckRow() As Integer = Nothing
        Dim ItemCountCheckContents() As Integer = Nothing
        Dim ItemCountCheckCount As Integer = 0
        Dim ItemCountFlg As Boolean = True
        Dim ItemCountErrorMessage As String = Nothing

        'DBのカラム名を格納する
        'DBのカラム数をカウントする。
        Dim DBColumnCount As Integer = 0

        Dim SQLWord As String = Nothing
        Dim SQLWord_Copy As String = Nothing
        Dim Word_Temp As String = Nothing

        '項目名が重複している場合の重複数を格納
        Dim GroupCount As Integer = 0
        'ファイル名を設定する際に使用
        Dim FileName As String = Nothing
        'ループ用変数
        Dim i As Integer = 0
        Dim j As Integer = 0

        'エラーデータを格納
        Dim ErrorMessageList() As String = Nothing
        'エラー件数をカウント
        Dim ErrorMessageCount As Integer = 0
        'エラーのあった行のデータを全て格納
        Dim ErrorData() As Object = Nothing
        'エラー行数を格納
        Dim ErrorDataCount() As Integer = Nothing
        'エラー項目名を格納
        Dim ErrorDataItem() As String = Nothing

        Dim ErrorData_Flg As Boolean = False

        '項目名の前にA_○○とつける為の変数。項目名のないケースでは、A_がそのまま項目名となる。（エクセルのようにZ列->AA列対応）
        Dim ItemName As String = Nothing

        '項目行がない場合、1行目のデータを一時的に格納する配列
        Dim Data() As String = Nothing
        '１行目に項目行がない場合のInsert文におけるループ回数を制御する変数
        Dim LoopControl As Integer = 0

        Dim ErrorDataNo() As Integer = Nothing
        'エラー行数を格納
        Dim ErrorDataNoCount As Integer = 0
        'Excelファイルを生成するかどうかを判別するフラグ
        Dim ExcelCreateFlg As Boolean = False
        'データを格納
        Dim DataList() As String = Nothing

        Dim DataListCount = 0

        Dim Loopcount As Integer = 0

        Dim EndMessage As String = Nothing
        'データ取り込み時の配列件数（読み込み時、ループのたびにRedimすると遅くなるのでとりあえずで設定）
        Dim SQLDataMax As Integer = 10000
        '配列をReDim。
        ReDim SQLData(0 To SQLDataMax)

        'データ連結する際の変数
        Dim EndCheck As String = Nothing

        ',オブジェクト変数の宣言
        Dim OBJ As Object

        '開くエクセルファイルのある、ディレクトリ、ファイル名を格納
        Dim fnm As String

        ' OpenFileDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
        Dim OpenFileDialog1 As New OpenFileDialog()

        ' ダイアログのタイトルを設定する
        OpenFileDialog1.Title = "ファイルを選択してください。"

        ' 初期表示するディレクトリを設定する
        OpenFileDialog1.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal)

        ' 初期表示するファイル名を設定する
        OpenFileDialog1.FileName = ""

        ' ファイルのフィルタを設定する
        'OpenFileDialog1.Filter = "xlsファイル|*.xls;"
        OpenFileDialog1.Filter = "csvファイル|*.csv;|xlsファイル|*.xls;|すべてのファイル|*.*"

        ' ファイルの種類 の初期設定を 2 番目に設定する (初期値 1)
        OpenFileDialog1.FilterIndex = 1

        ' ダイアログボックスを閉じる前に現在のディレクトリを復元する (初期値 False)
        OpenFileDialog1.RestoreDirectory = True

        ' 複数のファイルを選択可能にする (初期値 False)
        OpenFileDialog1.Multiselect = False

        ' [ヘルプ] ボタンを表示する (初期値 False)
        OpenFileDialog1.ShowHelp = False

        ' [読み取り専用] チェックボックスを表示する (初期値 False)
        OpenFileDialog1.ShowReadOnly = False

        ' [読み取り専用] チェックボックスをオンにする (初期値 False)
        OpenFileDialog1.ReadOnlyChecked = False

        ' 存在しないファイルを指定した場合は警告を表示する (初期値 True)
        'OpenFileDialog1.CheckFileExists = True

        ' 存在しないパスを指定した場合は警告を表示する (初期値 True)
        'OpenFileDialog1.CheckPathExists = True

        ' 拡張子を指定しない場合は自動的に拡張子を付加する (初期値 True)
        'OpenFileDialog1.AddExtension = True

        ' 有効な Win32 ファイル名だけを受け入れるようにする (初期値 True)
        'OpenFileDialog1.ValidateNames = True

        ' ダイアログを表示し、キャンセルボタンが押された場合、処理終了
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then
            Exit Sub
        End If

        'DataGridViewをクリアする。
        DataGridView1.Rows.Clear()

        '重複チェック用のcombbox1～5のクリア 2012/11/08対応
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox5.Items.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        ComboBox4.Text = ""
        ComboBox5.Text = ""

        'コンボボックスを使用不可能にする
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        ComboBox4.Enabled = False
        ComboBox5.Enabled = False
        ComboBox6.Enabled = False

        'オプション（住所結合１，２、会員メモ結合、反響メモ結合のcomboboxのクリア
        ComboBox7.Items.Clear()
        ComboBox8.Items.Clear()
        ComboBox9.Items.Clear()
        ComboBox10.Items.Clear()
        ComboBox11.Items.Clear()
        ComboBox12.Items.Clear()
        ComboBox13.Items.Clear()
        ComboBox14.Items.Clear()
        ComboBox15.Items.Clear()
        ComboBox16.Items.Clear()
        ComboBox17.Items.Clear()
        ComboBox18.Items.Clear()
        ComboBox19.Items.Clear()
        ComboBox20.Items.Clear()
        ComboBox21.Items.Clear()
        ComboBox22.Items.Clear()
        ComboBox23.Items.Clear()
        ComboBox24.Items.Clear()

        'AND検索にボタンを戻す。2012/11/08対応
        RadioButton3.Checked = True

        '不備・重複、取り込み可能リストの作成（店舗へのデータ確認用）にボタンを戻す。2012/11/08対応
        RadioButton1.Checked = True

        'mdbファイル名に使用する為、日付を取得
        Dim dtNow As DateTime = DateTime.Now
        ' 日付の部分だけを取得する
        Dim dtfilename As String = dtNow.ToString("yyyyMMddHHmmss")

        '選択されたファイルの拡張子がcsv or xlsかチェック
        Dim FileExtension As String = System.IO.Path.GetExtension(OpenFileDialog1.FileName)
        'もし拡張子が.csvじゃなかったらエラー終了
        If FileExtension <> ".csv" Then
            MsgBox("拡張子が.csvのファイルを指定してください。")
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        'ファイル名を表示
        TextBox1.Text = Path.GetFileName(OpenFileDialog1.FileName)
        'ディレクトリ名を取得
        Dim CSVDirName = Path.GetDirectoryName(OpenFileDialog1.FileName)
        '読み込んだファイル名をセット
        FileName = TextBox1.Text
        If FileExtension = ".csv" Then
            FileName = FileName.Replace(".csv", "")
        Else
            MsgBox("拡張子が.csvのファイルを指定してください。")
        End If

        'DBのパスを設定
        strMdbpath = CSVDirName + "\" + FileName + dtfilename + ".mdb"
        'もしパスが取得できていなければメッセージを表示し終了
        If strMdbpath = "" Then
            MsgBox("MDB作成ディレクトリを入力してください")
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        If Dir(strMdbpath) <> "" Or Dir(strMdbpath + ".mdb") <> "" Then
            'ファイルがあれば削除する。
            System.IO.File.Delete(strMdbpath)
        End If

        'ファイルを読み込む
        Dim parser As New TextFieldParser(CSVDirName & "\" & TextBox1.Text, System.Text.Encoding.GetEncoding("Shift_JIS"))
        '区切り文字の設定
        parser.TextFieldType = FieldType.Delimited
        parser.SetDelimiters(",")

        'データの終わりまでループ
        While Not parser.EndOfData
            ' 1行読み込み
            Dim row As String() = parser.ReadFields()

            '1行目の処理
            If DataCount = 0 Then
                '項目名が無い場合
                If CheckBox1.Checked = False Then
                    'DBのカラム、コンボボックスのリストの配列を再設定
                    ReDim Preserve DBColumnList(0 To row.Length - 1)
                    Loopcount = 0
                    'コンマ区切りのデータ分ループ
                    For Each field As String In row
                        DBColumnList(Loopcount) = field.Replace("'", "''")
                        'もし半角のカッコ ()　があれば、全角にする。
                        '半角カッコのままだと、MDBのカラム作成時にエラーになるが、名前（姓）でよく使用されている為、対応。
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("(", "（")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(")", "）")
                        'ピリオドがあれば全角に変換 2013/09/19 H.Mitao
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(".", "．")
                        '不等号、角かっこ、中かっこ、波かっこ、プラス、コロン、セミコロン、
                        'アポストロフィ、シャープ、パーセント、感嘆符、疑問符、アスタリスク、ドル記号
                        'があれば全角に変換 2013/09/20 H.Mitao
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("<", "＜")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(">", "＞")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("(", "（")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(")", "）")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("{", "｛")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("}", "｝")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("+", "＋")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(":", "：")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace(";", "；")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("'", "’")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("#", "＃")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("%", "％")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("!", "！")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("?", "？")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("*", "＊")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("$", "＄")
                        '[]があった場合も全角に変換 2013/10/18
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("[", "「")
                        DBColumnList(Loopcount) = DBColumnList(Loopcount).Replace("]", "」")
                        'ループ数から、A、Bなどを求める関数を呼ぶ
                        Result = ItemNameChk(Loopcount, ItemName, Result)
                        If Result = False Then
                            MsgBox("項目名の作成に失敗しました。")
                            OpenFileDialog1.Dispose()
                            Exit Sub
                        End If
                        DBColumnList(Loopcount) = ItemName & "_" & DBColumnList(Loopcount)
                        TitleItemCount += 1
                        Loopcount += 1
                        intCnt += 1
                    Next

                    ' C_の項目の作成
                    For i = 0 To row.Length - 1
                        ReDim Preserve DBColumnList(0 To i + row.Length)
                        DBColumnList(i + row.Length) = "C_" & DBColumnList(i)
                    Next
                Else
                    'チェックのある場合は１行目からデータとして、
                    '項目名はエクセルのようにA,B,Cと作成していき、読み込んだデータはDB作成後、Insertする。
                    ReDim Preserve Data(0 To 0)
                    Loopcount = 0
                    For Each field As String In row
                        'カラム数を１行目のデータでチェックする為、１行目のデータを配列に格納し後でInsertを行う。
                        If Loopcount = row.Length - 1 Then
                            EndCheck = "'"
                        Else
                            EndCheck = "',"
                        End If
                        SQLWord &= "'" & field.Replace("'", "''") & EndCheck
                        'コピー用は全角を半角にして、スペースも削除をして格納する。
                        Word_Temp = StrConv(field.Replace("'", "''"), VbStrConv.Narrow)
                        Word_Temp = Word_Temp.Replace("'", "''")
                        SQLWord_Copy &= "'" & Word_Temp.Replace(" ", "") & EndCheck
                        TitleItemCount += 1
                        Loopcount += 1
                        '項目数のカウントをループにあわせて+1する。
                        intCnt += 1
                    Next

                    Data(0) = SQLWord & "," & SQLWord_Copy

                    'C_の分の枠も作成するため配列再設定
                    ReDim Preserve DBColumnList(0 To (row.Length * 2) - 1)
                    'ループ数から、章と余りを算出し、A、B、Cとカラム名を生成していく。
                    For i = 0 To row.Length - 1
                        'ループ数から、A、Bなどを求める関数を呼ぶ
                        Result = ItemNameChk(i, ItemName, Result)
                        If Result = False Then
                            MsgBox("項目名の作成に失敗しました。")
                            OpenFileDialog1.Dispose()
                            Exit Sub
                        End If
                        DBColumnList(i) = ItemName & "_"
                    Next

                    ' C_の項目の作成
                    For i = 0 To row.Length - 1
                        DBColumnList(i + row.Length) = "C_" & DBColumnList(i)
                    Next
                End If

                DBColumnCount += 1

                '読み込み項目名*2（C_項目名の分）を代入
                ALL_ColumnCount = row.Length * 2

            ElseIf DataCount <> 0 Then
                '2行目以降の処理
                SQLWord = ""
                SQLWord_Copy = ""
                ErrorDataNoCount = 0
                ErrorData_Flg = False
                Loopcount = 0

                For Each field As String In row
                    If Loopcount = row.Length - 1 Then
                        EndCheck = "'"
                    Else
                        EndCheck = "',"
                    End If
                    ' 'を''にReplace
                    SQLWord &= "'" & field.Replace("'", "''") & EndCheck
                    Word_Temp = StrConv(field.Replace("'", "''"), VbStrConv.Narrow)
                    Word_Temp = Word_Temp.Replace("'", "''")
                    SQLWord_Copy &= "'" & Word_Temp.Replace(" ", "") & EndCheck

                    Loopcount += 1
                Next
                'もし、配列の現時点でのMAX件数をこえたら、配列を再設定する。
                If DataCount > SQLDataMax Then
                    SQLDataMax = SQLDataMax + 50000
                    '配列の再設定を行う。
                    ReDim Preserve SQLData(0 To SQLDataMax)
                End If
                'データを入れる。
                SQLData(DataCount - 1) = SQLWord & "," & SQLWord_Copy

                '項目数とデータ行数が違う場合、エラーメッセージ用配列に格納。
                If intCnt <> row.Length Then
                    ReDim Preserve ItemCountCheckRow(0 To ItemCountCheckCount)
                    ReDim Preserve ItemCountCheckContents(0 To ItemCountCheckCount)
                    ItemCountCheckRow(ItemCountCheckCount) = DataCount
                    ItemCountCheckContents(ItemCountCheckCount) = row.Length
                    ItemCountCheckCount += 1
                    ItemCountFlg = False
                End If
            End If

            DataCount += 1
        End While

        ReDim Preserve SQLData(0 To DataCount - 2)

        'DBの作成、テーブルを作成する。
        If MDB_CRTDATABASE(strMdbpath, DBColumnList) Then
        Else
            'mdbファイルを作成する際にエラーが出た場合、項目名が空白だったり、記号が含まれている可能性が高いので
            '項目名を確認すること。
            MsgBox("データベースファイルの作成に失敗しました。")
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        'DB作成後、１行目が項目行ではないデータの場合、
        '1件目のデータを登録する(項目数をカウントするために読み込んでしまった為)
        If CheckBox1.Checked = True Then
            LoopControl = 1
            'データの登録を行う
            Result = MDB_INSERT(LoopControl, strMdbpath, SQLColumn, Data)
            If Result = False Then
                MsgBox("データの登録に失敗しました。")
                OpenFileDialog1.Dispose()
                Exit Sub
            End If
        End If

        'フラグがFalseならメッセージを表示
        If ItemCountFlg = False Then
            For i = 0 To ItemCountCheckRow.Length - 1
                If i < 20 Then
                    ItemCountErrorMessage &= ItemCountCheckRow(i) & "行目（項目数：" & ItemCountCheckContents(i) & "）" & vbCr
                Else
                    EndMessage = "これ以上は表示できません。（エラー件数：" & ItemCountCheckRow.Length & "件）"
                End If
            Next
            MsgBox("以下の行数がタイトル項目数と異なっています（タイトル項目数：" & TitleItemCount & "）" & vbCr & ItemCountErrorMessage & vbCr & EndMessage)
            '破棄して終了
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        '何件目からの登録かによりPKが異なる為、LoopControlに値を入れる。
        If CheckBox1.Checked = True Then
            'Trueの場合、1行目からデータ行の為、プログラム上部ですでに1件登録済みの為、2件目からカウント
            LoopControl = 2
        Else
            'Falseの場合、1行目は項目行だった為、1件目として登録を行う。
            LoopControl = 1
        End If

        'データの登録を行う
        Result = MDB_INSERT(LoopControl, strMdbpath, SQLColumn, SQLData)
        If Result = False Then
            MsgBox("データの登録に失敗しました。")
            OpenFileDialog1.Dispose()
            Exit Sub
        End If

        '重複チェック項目を指定するコンボボックスに項目名を設定
        For i = 0 To DBColumnList.Length - 1
            ComboBox1.Items.Add(DBColumnList(i))
            ComboBox2.Items.Add(DBColumnList(i))
            ComboBox3.Items.Add(DBColumnList(i))
            ComboBox4.Items.Add(DBColumnList(i))
            ComboBox5.Items.Add(DBColumnList(i))
            '住所結合用(住所1)に以下のコンボボックスにも取り込んだ項目名を追加する。
            ComboBox7.Items.Add(DBColumnList(i))
            ComboBox8.Items.Add(DBColumnList(i))
            ComboBox9.Items.Add(DBColumnList(i))
            ComboBox10.Items.Add(DBColumnList(i))
            '住所結合用(住所2)に以下のコンボボックスにも取り込んだ項目名を追加する。
            ComboBox11.Items.Add(DBColumnList(i))
            ComboBox12.Items.Add(DBColumnList(i))
            ComboBox13.Items.Add(DBColumnList(i))
            ComboBox14.Items.Add(DBColumnList(i))
            '会員メモ用
            ComboBox15.Items.Add(DBColumnList(i))
            ComboBox16.Items.Add(DBColumnList(i))
            ComboBox17.Items.Add(DBColumnList(i))
            ComboBox18.Items.Add(DBColumnList(i))
            ComboBox19.Items.Add(DBColumnList(i))
            '反響メモ用
            ComboBox20.Items.Add(DBColumnList(i))
            ComboBox21.Items.Add(DBColumnList(i))
            ComboBox22.Items.Add(DBColumnList(i))
            ComboBox23.Items.Add(DBColumnList(i))
            ComboBox24.Items.Add(DBColumnList(i))
        Next

        '取得データ件数を表示
        MsgBox(SQLData.Length & "件のデータの取り込みが完了しました。")

        'コンボボックスを使用可能にする
        ComboBox1.Enabled = True
        ComboBox2.Enabled = True
        ComboBox3.Enabled = True
        ComboBox4.Enabled = True
        ComboBox5.Enabled = True
        ComboBox6.Enabled = True
        ComboBox7.Enabled = True
        ComboBox8.Enabled = True
        ComboBox9.Enabled = True
        ComboBox10.Enabled = True
        ComboBox11.Enabled = True
        ComboBox12.Enabled = True
        ComboBox13.Enabled = True
        ComboBox14.Enabled = True
        ComboBox15.Enabled = True
        ComboBox16.Enabled = True
        ComboBox17.Enabled = True
        ComboBox18.Enabled = True
        ComboBox19.Enabled = True
        ComboBox20.Enabled = True
        ComboBox21.Enabled = True
        ComboBox22.Enabled = True
        ComboBox23.Enabled = True
        ComboBox24.Enabled = True

        'データ作成ファイルを使用可能にする
        Button2.Enabled = True

        'ユーザ操作による行追加を無効(禁止)
        DataGridView1.AllowUserToAddRows = False

        'DataGridViewをクリアする
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        'DataGeidViewにカラムの登録
        For i = 0 To DBColumnList.Length - 1
            '列の追加
            DataGridView1.Columns.Add("clmName" & i, DBColumnList(i))

            '行の追加。
            If i = 0 Then
                DataGridView1.Rows.Add()
            End If

            'リストの設定
            Dim Category_Column As New DataGridViewComboBoxColumn()
            Category_Column.Items.Add("")
            For j = 0 To FormatList.Length - 1
                Category_Column.Items.Add(FormatList(j))
            Next

            'データを表示する
            Category_Column.DataPropertyName = DBColumnList(i)
            'ComboBox列を設定、表示
            DataGridView1.Columns.Insert(DataGridView1.Columns("clmName" & i).Index, Category_Column)
            DataGridView1.Columns.Remove("clmName" & i)
            Category_Column.Name = DBColumnList(i)
            DataGridView1.Columns(i).Width = 160

            'ドロップダウンリストの表示件数を10件にする
            Category_Column.MaxDropDownItems = 10
        Next

        '破棄する
        OpenFileDialog1.Dispose()
        'Dim span2 As TimeSpan = Now - start
        'MessageBox.Show(String.Format("処理時間は{0}秒です。", span2.TotalSeconds))
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'データチェック、ファイル作成ボタンが押された時
        Dim dtNow As DateTime
        '重複チェック項目を行うかどうかFalseなら行わない
        Dim Dupulicate_Check As Boolean = True
        Dim ErrorMessage As String = Nothing
        Dim Result As Boolean = True

        'ComboBox1～5のチェック内容を格納
        Dim Check1 As String
        Dim Check2 As String
        Dim Check3 As String
        Dim Check4 As String
        Dim Check5 As String

        '顧客全データを格納
        Dim MemberList As Object = Nothing

        '重複、不備のないデータ（取り込みデータ）
        Dim NormalDataList As Object = Nothing
        '取り込みデータ件数
        Dim NormalDataCount As Integer = 0

        'エラーのカウント
        Dim ErrorDataNoCount As Integer = 0
        'エラーデータの有無を判断
        Dim ErrorData_Flg As Boolean = True
        'エラーデータの場合、データを空白にしてDD用フォーマットに取り込むかを判断
        'ただし、郵便番号１、郵便番号２、PCメールアドレス、携帯メールアドレスの際にのみ有効
        Dim ErrorData_ImportFLg As Boolean = False
        '上記の４項目以外でエラーがあるかを判断。
        Dim ErrorData_ImporChecktFLg As Boolean = False

        'エラーのあった行のデータを全て格納
        Dim ErrorData() As Object = Nothing
        'エラー件数をカウント
        Dim ErrorMessageCount As Integer = 0
        '戻り配列
        Dim ErrorDataList() As String = Nothing

        'エラー項目名を格納
        Dim ErrorDataItem() As String = Nothing

        'Excel作成フラグ(True:作成、False:作成しない）
        Dim ExcelCreateFlg As Boolean = False

        Dim DataGridList() As String = Nothing

        Dim DataList() As String = Nothing

        '空白にして取り込むリストを格納
        Dim ErrorImportList() As Object = Nothing
        '取り込みデータ件数
        Dim ErrorImportDataCount As Integer = 0
        'エラーデータの行数を格納
        Dim ErrorImportDataNo() As String = Nothing
        'エラーデータの件数を格納
        Dim ErrorImportDataNoCount As Integer = 0
        'エラーの項目名を格納
        Dim ErrorImportDataContent() As String = Nothing

        Dim ErrorImportDataList() As String = Nothing

        'エラーのあった行のデータを全て格納
        Dim ErrorImportData() As Object = Nothing
        'エラー件数をカウント
        Dim ErrorImportMessageCount As Integer = 0
        'エラー項目名を格納
        Dim ErrorImportDataItem() As String = Nothing

        Dim ErrorImportDataItemName() As String = Nothing

        Dim ExcelImportCreateFlg As Boolean = False

        '住所1結合用変数
        Dim lp1 As Integer = 0
        Dim address1data1 As String = Nothing
        Dim address1data2 As String = Nothing
        Dim address1data3 As String = Nothing
        Dim address1data4 As String = Nothing
        Dim tmp1_1 As Integer = 0
        Dim tmp1_2 As Integer = 0
        Dim tmp1_3 As Integer = 0
        Dim tmp1_4 As Integer = 0
        Dim tmp1address As String = Nothing
        Dim AdressUnion1_flg As Boolean = False

        '住所2結合用変数
        Dim lp2 As Integer = 0
        Dim address2data1 As String = Nothing
        Dim address2data2 As String = Nothing
        Dim address2data3 As String = Nothing
        Dim address2data4 As String = Nothing
        Dim tmp2_1 As Integer = 0
        Dim tmp2_2 As Integer = 0
        Dim tmp2_3 As Integer = 0
        Dim tmp2_4 As Integer = 0
        Dim tmp2address As String = Nothing
        Dim AdressUnion2_flg As Boolean = False

        'ループ用変数
        Dim lp As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0
        Dim y As Integer = 0

        '会員メモ結合機能に使用
        Dim Tmp_Memo As String = Nothing
        Dim Tmp_Memo1 As Integer = 0
        Dim Tmp_Memo2 As Integer = 0
        Dim Tmp_Memo3 As Integer = 0
        Dim Tmp_Memo4 As Integer = 0
        Dim MemoUnion_flg As Boolean = False

        '反響メモ結合機能に使用
        Dim Tmp_Echo As String = Nothing
        Dim Tmp_Echo1 As Integer = 0
        Dim Tmp_Echo2 As Integer = 0
        Dim Tmp_Echo3 As Integer = 0
        Dim Tmp_Echo4 As Integer = 0
        Dim EchoUnion_flg As Boolean = False

        '姓名分割後データ格納配列
        Dim TmpData As String() = Nothing

        '取り込んだCSVファイル名を格納
        Dim FileName As String = Nothing

        '郵便番号チェックのリターン結果を格納する変数
        Dim ReturnString As String = Nothing

        Dim TmpPost1 As String = Nothing
        Dim TmpPost2 As String = Nothing
        Dim TmpPCaddress As String = Nothing
        Dim TmpMobileaddress As String = Nothing

        'カラム名を格納
        Dim TmpErrorImportColData() As String = Nothing
        '実データが格納
        Dim TmpErrorImportData() As String = Nothing

        '不備データチェックを行うかどうか。Falseなら行わない旨のアラートを表示
        Dim CheckSelectFlg As Boolean = False

        ErrorDataItemName = Nothing

        'Excel作成時のデータ貼り付け用始点設定。
        Dim startX As Integer = 1
        Dim startY As Integer = 1

        '不備・重複データのファイル名（ファイル名の後ろのかっこの中の記載+拡張子）
        Dim FileType1 As String = "（不備、重複データ）.xls"
        '取り込み可能データのファイル名（ファイル名の後ろのかっこの中の記載+拡張子）
        Dim FileType2 As String = "（取り込みデータ）.xls"
        'DD取り込み用フォーマットのファイル名（ファイル名の後ろのかっこの中の記載+拡張子）
        Dim FileType3 As String = "（DD取り込み用フォーマットデータ）.xls"
        'DD取り込み用フォーマットの一部データを空白にして取り込まれたデータの報告書ファイル名（ファイル名の後ろのかっこの中の記載+拡張子）
        Dim FileType4 As String = "（DD追加取り込み報告データ）.xls"

        'ANDかOR検索か
        Dim Search_Type As String = Nothing

        If RadioButton3.Checked = True Then
            Search_Type = "AND検索"
        Else
            Search_Type = "OR検索"
        End If

        ' ファイルが存在しているかどうか確認する
        If System.IO.File.Exists(strMdbpath) Then
        Else
            MessageBox.Show("DBファイルが存在しない為、処理を中止します。")
            Application.Exit()
        End If

        '同じ項目が指定されていればメッセージを表示し処理終了
        If ComboBox2.Text <> "" Then
            If ComboBox1.Text = ComboBox2.Text Or ComboBox3.Text = ComboBox2.Text Or _
               ComboBox4.Text = ComboBox2.Text Or ComboBox5.Text = ComboBox2.Text Then
                MsgBox("同じ項目を複数指定することはできません。")
                Exit Sub
            End If
        End If
        '同じ項目が指定されていればメッセージを表示し処理終了
        If ComboBox3.Text <> "" Then
            If ComboBox1.Text = ComboBox3.Text Or ComboBox2.Text = ComboBox3.Text Or _
               ComboBox4.Text = ComboBox3.Text Or ComboBox5.Text = ComboBox3.Text Then
                MsgBox("同じ項目を複数指定することはできません。")
                Exit Sub
            End If
        End If
        '2番目の項目が未設定で3番目の項目が指定されていたらメッセージを表示し処理終了
        If ComboBox3.Text <> "" And ComboBox2.Text = "" Then
            MsgBox("2番目の項目が未設定です。")
            Exit Sub
        End If
        '3番目の項目が未設定で4番目の項目が指定されていたらメッセージを表示し処理終了
        If ComboBox4.Text <> "" And ComboBox3.Text = "" Then
            MsgBox("3番目の項目が未設定です。")
            Exit Sub
        End If
        '4番目の項目が未設定で5番目の項目が指定されていたらメッセージを表示し処理終了
        If ComboBox5.Text <> "" And ComboBox4.Text = "" Then
            MsgBox("4番目の項目が未設定です。")
            Exit Sub
        End If

        '住所結合機能の指定内容確認（同じ項目を指定しているとメッセージを表示し終了
        If ComboBox8.Text <> "" Then
            If ComboBox7.Text = ComboBox8.Text Or ComboBox9.Text = ComboBox8.Text Or _
               ComboBox10.Text = ComboBox8.Text Then
                MsgBox("住所結合機能（住所1用）で同じ項目を複数指定することはできません。")
                Exit Sub
            End If
        End If
        '同じ項目が指定されていればメッセージを表示し処理終了
        If ComboBox9.Text <> "" Then
            If ComboBox7.Text = ComboBox9.Text Or ComboBox8.Text = ComboBox9.Text Or _
               ComboBox10.Text = ComboBox9.Text Then
                MsgBox("住所結合機能（住所1用）で同じ項目を複数指定することはできません。")
                Exit Sub
            End If
        End If

        '住所１用のチェック
        If ComboBox9.Text <> "" And ComboBox8.Text = "" Then
            MsgBox("住所結合機能（住所1用）の2番目の項目が未設定です。")
            Exit Sub
        End If
        '住所１用の3番目の項目が未設定で4番目の項目が指定されていたらメッセージを表示し処理終了
        If ComboBox10.Text <> "" And ComboBox9.Text = "" Then
            MsgBox("住所結合機能（住所1用）の3番目の項目が未設定です。")
            Exit Sub
        End If

        '重複チェック項目が指定されていなければアラートを表示。
        If ComboBox1.Text = "" And ComboBox2.Text = "" And ComboBox3.Text = "" And ComboBox4.Text = "" And ComboBox5.Text = "" Then
            'ダイアログ設定
            Dim check As DialogResult = MessageBox.Show("重複チェック項目が指定されていませんが、よろしいですか？", _
                                                         "確認", _
                                                         MessageBoxButtons.YesNo, _
                                                         MessageBoxIcon.Question)

            If check = DialogResult.No Then
                'ダイアログで「いいえ」が選択された時 
                Exit Sub
            Else
                Dupulicate_Check = False
            End If
        End If

        'データチェックを行う項目があるか確認
        For i = 0 To DataGridView1.ColumnCount - 1
            ReDim Preserve DataGridList(0 To i)
            If DataGridView1(i, 0).Value <> "" Then
                CheckSelectFlg = True
            End If
            DataGridList(i) = DataGridView1(i, 0).Value
        Next

        '不備データチェックを何も行わないならアラートを表示。
        If CheckSelectFlg = False Then

            'ダイアログ設定
            Dim check As DialogResult = MessageBox.Show("データの妥当性確認項目が設定されていませんが、よろしいですか？", _
                                                         "確認", _
                                                         MessageBoxButtons.YesNo, _
                                                         MessageBoxIcon.Question)
            If check = DialogResult.No Then
                'ダイアログで「いいえ」が選択された時 
                Exit Sub
            End If
        End If

        'どのデータの重複チェックをするか確認
        Check1 = ComboBox1.Text
        Check2 = ComboBox2.Text
        Check3 = ComboBox3.Text
        Check4 = ComboBox4.Text
        Check5 = ComboBox5.Text

        '処理中ウインドウを表示
        Processing.Show()

        'Processingウインドウの「処理中です。」のメッセージをリフレッシュさせることにより表示
        Processing.Label1.Refresh()

        '処理前にmdbの重複Noカラムのクリアを行う。 2012/11/08対応
        Result = MDB_DuplicateClear(strMdbpath, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            '処理中ウインドウを破棄
            Processing.Dispose()
            Exit Sub
        End If

        '重複チェックを行う。
        'Dupulicate_CheckがFlaseなら行わない。
        If Dupulicate_Check = True Then
            '最初は重複しているものを抽出し、そのデータに対し、テーブルの重複Noのカラムに重複しているPKを入れる。
            Result = MDB_DuplicateCheck(strMdbpath, Check1, Check2, Check3, Check4, Check5, Search_Type, Result, ErrorMessage)
            If Result = False Then
                MsgBox(ErrorMessage)
                '処理中ウインドウを破棄
                Processing.Dispose()
                Exit Sub
            End If
        End If

        '全データを取得し、DataGridで選択された項目のエラーチェックを行う。
        'データ取得
        Result = MDB_GET_MemberData(strMdbpath, MemberList, Result, ErrorMessage)
        If Result = False Then
            MsgBox(ErrorMessage)
            Exit Sub
        End If

        'DataGridで設定された項目とチェック内容を格納
        'データ件数分ループ
        For i = 0 To MemberList.Length - 1
            ErrorDataNoCount = 0
            ErrorData_Flg = False
            AdressUnion1_flg = False
            AdressUnion2_flg = False
            ErrorData_ImportFLg = False
            ErrorData_ImporChecktFLg = False
            MemoUnion_flg = False
            EchoUnion_flg = False

            ReDim TmpErrorImportData(0 To 0)
            ErrorImportDataNoCount = 0

            '項目数分ループ
            For j = 0 To DataGridList.Length - 1
                If DataGridList(j) = "登録店舗名" Then
                    ''登録店舗名の場合、文字列が20文字以内かチェックをする。
                    'If Trim(MemberList(i)(j + 1).Length) > 20 Or Trim(MemberList(i)(j + 1)) = "" Then
                    '    'エラー内容を格納する。
                    '    ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                    '    ErrorDataNo(ErrorDataNoCount) = j
                    '    ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                    '    ErrorDataContent(ErrorDataNoCount) = "登録店舗名"
                    '    ErrorDataNoCount += 1
                    '    ErrorData_Flg = True
                    'Else
                    '前後スペースを省いて格納する。
                    MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                    'End If
                ElseIf DataGridList(j) = "会員番号" Then
                    ReturnString = Nothing

                    '全角から半角に変換する。(MemberNoDataCheck関数内に移動。2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                    If (MemberNoDataCheck(Trim(MemberList(i)(j + 1)), ReturnString) = False) Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "会員番号"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True
                        ErrorData_ImporChecktFLg = True
                    Else
                        MemberList(i)(j + 1) = ReturnString
                    End If
                ElseIf DataGridList(j) = "顧客名(姓)" Then
                    '顧客名（姓）の場合、文字列が20文字以内かチェックをする。
                    If Trim(MemberList(i)(j + 1).Length) > 20 Or Trim(MemberList(i)(j + 1)) = "" Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "顧客名(姓)"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True
                        ErrorData_ImporChecktFLg = True
                    Else
                        '前後スペースを省い格納する。
                        MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                    End If
                ElseIf DataGridList(j) = "顧客名(名)" Then
                    '顧客名（名）の場合、文字列が20文字以内かチェックをする。
                    If Trim(MemberList(i)(j + 1).Length) > 20 Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "顧客名(名)"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True
                        ErrorData_ImporChecktFLg = True
                    Else
                        '前後スペースを省いて格納する。
                        MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                    End If
                ElseIf DataGridList(j) = "ふりがな(姓)" Then
                    'ふりがな（姓）の場合、文字列が20文字以内かチェックをする。
                    If Trim(MemberList(i)(j + 1).Length) > 20 Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "ふりがな(姓)"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True
                        ErrorData_ImporChecktFLg = True
                    Else
                        '前後スペースを省いて格納する。
                        MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                    End If
                ElseIf DataGridList(j) = "ふりがな(名)" Then
                    'ふりがな（名）の場合、文字列が20文字以内かチェックをする。
                    If Trim(MemberList(i)(j + 1).Length) > 20 Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "ふりがな(名)"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True
                        ErrorData_ImporChecktFLg = True
                    Else
                        '前後スペースを省いて格納する。
                        MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                    End If
                ElseIf DataGridList(j) = "郵便番号1" Then
                    ReturnString = Nothing
                    '全角から半角に変換する。(PostDataCheck関数内に移動。2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                    If (PostDataCheck(Trim(MemberList(i)(j + 1)), True, ReturnString) = False) Then
                        'エラー内容を格納する。
                        ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                        ErrorDataNo(ErrorDataNoCount) = j
                        ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                        ErrorDataContent(ErrorDataNoCount) = "郵便番号1"
                        ErrorDataNoCount += 1
                        ErrorData_Flg = True

                        '空白にして取り込む用にデータを格納。設定。
                        'DD向けのファイル作成なら。
                        If RadioButton2.Checked = True Then
                            ErrorData_ImportFLg = True

                            ReDim Preserve ErrorImportDataNo(0 To ErrorImportDataNoCount)
                            ErrorImportDataNo(ErrorImportDataNoCount) = j
                            'DD向けファイル生成時のみ、空白にして取り込む為、""を入れる。

                            ReDim Preserve TmpErrorImportData(0 To ErrorImportDataNoCount)
                            TmpErrorImportData(ErrorImportDataNoCount) = MemberList(i)(j + 1)
                            MemberList(i)(j + 1) = ""

                            ReDim Preserve ErrorImportDataContent(0 To ErrorImportDataNoCount)
                            ErrorImportDataContent(ErrorImportDataNoCount) = "郵便番号1"
                            ErrorImportDataNoCount += 1
                        End If
                    Else
                        MemberList(i)(j + 1) = ReturnString
                    End If
                    ElseIf DataGridList(j) = "住所1" And AdressUnion1_flg = False Then
                        '住所結合機能が指定されていれば、指定の順番で結合を行う。
                        If ComboBox7.Text <> "" Then
                            For lp1 = 0 To DataGridList.Length - 1
                                If DataGridList(lp1) = "住所1" And DBColumnList(lp1) = ComboBox7.Text Then
                                    address1data1 = Trim(MemberList(i)(lp1 + 1))
                                    tmp1_1 = lp1
                                    Exit For
                                End If
                            Next
                            If ComboBox8.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所1" And tmp1_1 <> lp1 And DBColumnList(lp1) = ComboBox8.Text Then
                                        address1data2 = Trim(MemberList(i)(lp1 + 1))
                                        tmp1_2 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox9.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所1" And tmp1_1 <> lp1 And tmp1_2 <> lp1 And DBColumnList(lp1) = ComboBox9.Text Then
                                        address1data3 = Trim(MemberList(i)(lp1 + 1))
                                        tmp1_3 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox10.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所1" And tmp1_1 <> lp1 And tmp1_2 <> lp1 And tmp1_3 <> lp1 And DBColumnList(lp1) = ComboBox10.Text Then
                                        address1data4 = Trim(MemberList(i)(lp1 + 1))
                                        Exit For
                                    End If
                                Next
                            End If
                            tmp1address = address1data1 & address1data2 & address1data3 & address1data4
                            AdressUnion1_flg = True
                        Else
                            If AdressUnion1_flg = True And (DBColumnList(j) = ComboBox8.Text Or DBColumnList(j) = ComboBox9.Text Or DBColumnList(j) = ComboBox10.Text) Then
                            Else
                                tmp1address = Trim(MemberList(i)(j + 1))
                            End If
                        End If
                        '住所1の場合、文字列が200文字以内かチェックをする。
                        If tmp1address.Length > 200 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "住所1"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            If RadioButton1.Checked = True Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                If AdressUnion1_flg = True Then
                                    MemberList(i)(tmp1_1 + 1) = tmp1address
                                Else
                                    MemberList(i)(j + 1) = tmp1address
                                End If
                            End If
                        End If
                    ElseIf DataGridList(j) = "郵便番号2" Then
                    ReturnString = Nothing
                    '全角から半角に変換する。(PostDataCheck関数内に移動。2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(Trim(MemberList(i)(j + 1)), VbStrConv.Narrow)
                        If (PostDataCheck(MemberList(i)(j + 1), True, ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "郵便番号2"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True

                            '空白にして取り込む用にデータを格納。設定。
                            'DD向けのファイル作成なら。
                            If RadioButton2.Checked = True Then
                                ErrorData_ImportFLg = True

                                ReDim Preserve ErrorImportDataNo(0 To ErrorImportDataNoCount)
                                ErrorImportDataNo(ErrorImportDataNoCount) = j
                                'DD向けファイル生成時のみ、空白にして取り込む為、""を入れる。

                                ReDim Preserve TmpErrorImportData(0 To ErrorImportDataNoCount)
                                TmpErrorImportData(ErrorImportDataNoCount) = MemberList(i)(j + 1)

                                MemberList(i)(j + 1) = ""

                                ReDim Preserve ErrorImportDataContent(0 To ErrorImportDataNoCount)
                                ErrorImportDataContent(ErrorImportDataNoCount) = "郵便番号2"
                                ErrorImportDataNoCount += 1
                            End If
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "住所2" And AdressUnion2_flg = False Then
                        If ComboBox11.Text <> "" Then
                            For lp1 = 0 To DataGridList.Length - 1
                                If DataGridList(lp1) = "住所2" And DBColumnList(lp1) = ComboBox11.Text Then
                                    address2data1 = Trim(MemberList(i)(lp1 + 1))
                                    tmp2_1 = lp1
                                    Exit For
                                End If
                            Next
                            If ComboBox12.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所2" And tmp2_1 <> lp1 And DBColumnList(lp1) = ComboBox12.Text Then
                                        address2data2 = Trim(MemberList(i)(lp1 + 1))
                                        tmp2_2 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox13.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所2" And tmp2_1 <> lp1 And tmp2_2 <> lp1 And DBColumnList(lp1) = ComboBox13.Text Then
                                        address2data3 = Trim(MemberList(i)(lp1 + 1))
                                        tmp2_3 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox14.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "住所2" And tmp2_1 <> lp1 And tmp2_2 <> lp1 And tmp2_3 <> lp1 And DBColumnList(lp1) = ComboBox14.Text Then
                                        address2data4 = Trim(MemberList(i)(lp1 + 1))
                                        Exit For
                                    End If
                                Next
                            End If
                            tmp2address = address2data1 & address2data2 & address2data3 & address2data4
                            AdressUnion2_flg = True
                        Else
                            If AdressUnion2_flg = True And (DBColumnList(j) = ComboBox12.Text Or DBColumnList(j) = ComboBox13.Text Or DBColumnList(j) = ComboBox14.Text) Then
                            Else
                                tmp2address = Trim(MemberList(i)(j + 1))
                            End If
                        End If

                        '住所2の場合、文字列が200文字以内かチェックをする。
                        If Trim(tmp2address.Length) > 200 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "住所2"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            If RadioButton1.Checked = True Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                If AdressUnion1_flg = True Then
                                    MemberList(i)(tmp2_1 + 1) = tmp2address
                                Else
                                    MemberList(i)(j + 1) = tmp2address
                                End If
                            End If
                        End If
                    ElseIf DataGridList(j) = "電話番号(自宅)" Then
                        ReturnString = Nothing
                    'ハイフンを取り除く(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = MemberList(i)(j + 1).Replace("-", "")
                    '全角から半角に変換する。(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                        If (TelDataCheck(Trim(MemberList(i)(j + 1)), True, ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "電話番号(自宅)"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "電話番号(携帯)" Then
                        ReturnString = Nothing
                    ''ハイフンを取り除く(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = MemberList(i)(j + 1).Replace("-", "")
                    ''全角から半角に変換する。(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                        If (TelDataCheck(Trim(MemberList(i)(j + 1)), True, ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "電話番号(携帯)"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "電話番号(会社)" Then
                        ReturnString = Nothing
                    ''ハイフンを取り除く(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = MemberList(i)(j + 1).Replace("-", "")
                    ''全角から半角に変換する。(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                        If (TelDataCheck(Trim(MemberList(i)(j + 1)), True, ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "電話番号(会社)"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "電話番号(その他)" Then
                        ReturnString = Nothing
                    ''ハイフンを取り除く(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = MemberList(i)(j + 1).Replace("-", "")
                    ''全角から半角に変換する。(TelDataCheck関数内に移動 2013/09/20 H.Mitao)
                    'MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                        If (TelDataCheck(Trim(MemberList(i)(j + 1)), True, ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "電話番号(その他)"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "性別" Then
                        '0=女性、1=男性。空白または該当する値がない場合は女性として設定
                        If Trim(MemberList(i)(j + 1)) = "1" Then
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        Else
                            '空白や、該当のない値が入っている可能性があるので0を設定する。
                            MemberList(i)(j + 1) = 0
                        End If
                    ElseIf DataGridList(j) <> "" And DataGridList(j) = "誕生日" Then
                        '誕生日の場合、日付としての妥当性チェックを行う。

                        '全角から半角に変換する。
                        MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)
                        'nullならチェックを行わない
                        If Trim(MemberList(i)(j + 1)) <> "" Then
                            If IsDate(Trim(MemberList(i)(j + 1))) Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                'エラー内容を格納する。
                                ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                                ErrorDataNo(ErrorDataNoCount) = j
                                ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                                ErrorDataContent(ErrorDataNoCount) = "誕生日"
                                ErrorDataNoCount += 1
                                ErrorData_Flg = True
                                ErrorData_ImporChecktFLg = True
                            End If
                        End If
                    ElseIf DataGridList(j) = "初回来店日" Then
                        'nullならチェックを行わない
                        If Trim(MemberList(i)(j + 1)) <> "" Then
                            '初回来店日の場合、日付としての妥当性チェックを行う。
                            'ただし、YYYY/MM/DD形式のみOKとする。
                        'If System.Text.RegularExpressions.Regex.IsMatch(Trim(MemberList(i)(j + 1)), "\d{4}/\d{1,2}/\d{1,2}", _
                        'System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                        If IsDate(Trim(MemberList(i)(j + 1))) Then
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        Else
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "初回来店日"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        End If
                        End If
                    ElseIf DataGridList(j) = "PCﾒｰﾙｱﾄﾞﾚｽ" Then
                        ReturnString = Nothing
                        If (MailAddressNoDataCheck(Trim(MemberList(i)(j + 1)), ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "PCﾒｰﾙｱﾄﾞﾚｽ"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True

                            '空白にして取り込む用にデータを格納。設定。
                            'DD向けのファイル作成なら。
                            If RadioButton2.Checked = True Then
                                ErrorData_ImportFLg = True

                                ReDim Preserve ErrorImportDataNo(0 To ErrorImportDataNoCount)
                                ErrorImportDataNo(ErrorImportDataNoCount) = j
                                'DD向けファイル生成時のみ、空白にして取り込む為、""を入れる。

                                ReDim Preserve TmpErrorImportData(0 To ErrorImportDataNoCount)
                                TmpErrorImportData(ErrorImportDataNoCount) = MemberList(i)(j + 1)

                                MemberList(i)(j + 1) = ""

                                ReDim Preserve ErrorImportDataContent(0 To ErrorImportDataNoCount)
                                ErrorImportDataContent(ErrorImportDataNoCount) = "PCﾒｰﾙｱﾄﾞﾚｽ"
                                ErrorImportDataNoCount += 1
                            End If
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "PCﾒｰﾙ受信許可" Then
                        '0=不許可、1=許可。空白は不許可として登録。
                        If Trim(MemberList(i)(j + 1)) = "1" Then
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        ElseIf Trim(MemberList(i)(j + 1)) = "0" Or Trim(MemberList(i)(j + 1)) = "" Then
                            '0か空白なら0を設定する。
                            MemberList(i)(j + 1) = 0
                        Else
                            'それ以外のものは入っていたらエラー。
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "PCﾒｰﾙ受信許可"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = False
                        End If
                    ElseIf DataGridList(j) = "携帯ﾒｰﾙｱﾄﾞﾚｽ" Then
                        ReturnString = Nothing
                        If (MailAddressNoDataCheck(Trim(MemberList(i)(j + 1)), ReturnString) = False) Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "携帯ﾒｰﾙｱﾄﾞﾚｽ"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True

                            '空白にして取り込む用にデータを格納。設定。
                            'DD向けのファイル作成なら。
                            If RadioButton2.Checked = True Then
                                ErrorData_ImportFLg = True

                                ReDim Preserve ErrorImportDataNo(0 To ErrorImportDataNoCount)
                                ErrorImportDataNo(ErrorImportDataNoCount) = j
                                'DD向けファイル生成時のみ、空白にして取り込む為、""を入れる。

                                ReDim Preserve TmpErrorImportData(0 To ErrorImportDataNoCount)
                                TmpErrorImportData(ErrorImportDataNoCount) = MemberList(i)(j + 1)

                                MemberList(i)(j + 1) = ""

                                ReDim Preserve ErrorImportDataContent(0 To ErrorImportDataNoCount)
                                ErrorImportDataContent(ErrorImportDataNoCount) = "携帯ﾒｰﾙｱﾄﾞﾚｽ"
                                ErrorImportDataNoCount += 1
                            End If
                        Else
                            MemberList(i)(j + 1) = ReturnString
                        End If
                    ElseIf DataGridList(j) = "携帯ﾒｰﾙ受信許可" Then
                        '0=不許可、1=許可。空白は不許可として登録。
                        If Trim(MemberList(i)(j + 1)) = "1" Then
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        ElseIf Trim(MemberList(i)(j + 1)) = "0" Or Trim(MemberList(i)(j + 1)) = "" Then
                            '0か空白なら0を設定する。
                            MemberList(i)(j + 1) = 0
                        Else
                            'それ以外のものは入っていたらエラー。
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "携帯ﾒｰﾙ受信許可"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        End If
                    ElseIf DataGridList(j) = "主担当" Then
                        '主担当の場合、文字列20文字以内かチェック
                        If Trim(MemberList(i)(j + 1).Length) > 20 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "主担当"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        End If
                    ElseIf DataGridList(j) = "DM許可フラグ" Then
                        '0=不許可、1=許可。空白は不許可として登録。
                        If Trim(MemberList(i)(j + 1)) = "1" Then
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        ElseIf Trim(MemberList(i)(j + 1)) = "0" Or Trim(MemberList(i)(j + 1)) = "" Then
                            '0か空白なら0を設定する。
                            MemberList(i)(j + 1) = 0
                        Else
                            'それ以外のものは入っていたらエラー。
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "DM許可フラグ"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        End If
                    ElseIf DataGridList(j) = "整理番号" Then
                        '整理番号の場合、文字列20文字以内かチェック
                        If Trim(MemberList(i)(j + 1).Length) > 20 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "整理番号"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        End If
                    ElseIf DataGridList(j) = "会員メモ" And MemoUnion_flg = False Then
                        '会員メモ結合機能が設定されていたら
                        Tmp_Memo = ""
                        If ComboBox15.Text <> "" Then
                            For lp1 = 0 To DataGridList.Length - 1
                                If DataGridList(lp1) = "会員メモ" And DBColumnList(lp1) = ComboBox15.Text Then
                                    Tmp_Memo = Trim(MemberList(i)(lp1 + 1))
                                    Tmp_Memo1 = lp1
                                    Exit For
                                End If
                            Next
                            If ComboBox16.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "会員メモ" And Tmp_Memo1 <> lp1 And DBColumnList(lp1) = ComboBox16.Text Then
                                        Tmp_Memo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Memo2 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox17.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "会員メモ" And Tmp_Memo1 <> lp1 And Tmp_Memo2 <> lp1 And DBColumnList(lp1) = ComboBox17.Text Then
                                        Tmp_Memo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Memo3 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox18.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "会員メモ" And Tmp_Memo1 <> lp1 And Tmp_Memo2 <> lp1 And Tmp_Memo3 <> lp1 And DBColumnList(lp1) = ComboBox18.Text Then
                                        Tmp_Memo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Memo4 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox19.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "会員メモ" And Tmp_Memo1 <> lp1 And Tmp_Memo2 <> lp1 And Tmp_Memo3 <> lp1 And Tmp_Memo4 <> lp1 And DBColumnList(lp1) = ComboBox19.Text Then
                                        Tmp_Memo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Exit For
                                    End If
                                Next
                            End If
                            MemoUnion_flg = True
                        Else
                            If MemoUnion_flg = True And (DBColumnList(j) = ComboBox16.Text Or DBColumnList(j) = ComboBox17.Text Or DBColumnList(j) = ComboBox18.Text Or DBColumnList(j) = ComboBox19.Text) Then
                            Else
                                Tmp_Memo = Trim(MemberList(i)(j + 1))
                            End If
                        End If

                        '会員メモの場合、文字列100文字以内かチェック
                        If Trim(Tmp_Memo.Length) > 100 Then
                            'エラー内容を格納する。
                            'ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            'ErrorDataNo(ErrorDataNoCount) = j
                            'ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            'ErrorDataContent(ErrorDataNoCount) = "会員メモ"
                            'ErrorDataNoCount += 1
                            'ErrorData_Flg = True

                            '101文字以上なら、1～100文字目まで切り取る。
                            MemberList(i)(j + 1) = Tmp_Memo.Substring(0, 101)
                        Else
                            '作成するファイルにより格納値を変える。
                            If RadioButton1.Checked = True Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                If MemoUnion_flg = True Then
                                    MemberList(i)(Tmp_Memo1 + 1) = Tmp_Memo
                                Else
                                    MemberList(i)(j + 1) = Tmp_Memo
                                End If
                            End If
                        End If
                    ElseIf DataGridList(j) = "反響日" Then
                        '反響日の場合、日付としての妥当性チェックを行う。
                        'nullならチェックを行わない
                        If Trim(MemberList(i)(j + 1)) <> "" Then
                            'ただし、YYYY/MM/DD形式のみOKとする。
                            If System.Text.RegularExpressions.Regex.IsMatch(Trim(MemberList(i)(j + 1)), "\d{4}/\d{1,2}/\d{1,2}", _
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                'エラー内容を格納する。
                                ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                                ErrorDataNo(ErrorDataNoCount) = j
                                ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                                ErrorDataContent(ErrorDataNoCount) = "反響日"
                                ErrorDataNoCount += 1
                                ErrorData_Flg = True
                                ErrorData_ImporChecktFLg = True
                            End If
                        End If
                    ElseIf DataGridList(j) = "媒体名" Then
                        '媒体名の場合、文字列20文字以内かチェック
                        If Trim(MemberList(i)(j + 1).Length) > 20 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "媒体名"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        End If
                    ElseIf DataGridList(j) = "反響電話担当者" Then
                        '反響電話担当者の場合、文字列20文字以内かチェック
                        If Trim(MemberList(i)(j + 1).Length) > 20 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "反響電話担当者"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        End If
                    ElseIf DataGridList(j) = "反響メモ" And EchoUnion_flg = False Then
                        '会員メモ結合機能が設定されていたら
                        Tmp_Echo = ""
                        If ComboBox20.Text <> "" Then
                            For lp1 = 0 To DataGridList.Length - 1
                                If DataGridList(lp1) = "反響メモ" And DBColumnList(lp1) = ComboBox20.Text Then
                                    Tmp_Echo = Trim(MemberList(i)(lp1 + 1))
                                    Tmp_Echo1 = lp1
                                    Exit For
                                End If
                            Next
                            If ComboBox21.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "反響メモ" And Tmp_Echo1 <> lp1 And DBColumnList(lp1) = ComboBox21.Text Then
                                        Tmp_Echo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Echo2 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox22.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "反響メモ" And Tmp_Echo1 <> lp1 And Tmp_Echo2 <> lp1 And DBColumnList(lp1) = ComboBox22.Text Then
                                        Tmp_Echo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Echo3 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox23.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "反響メモ" And Tmp_Echo1 <> lp1 And Tmp_Echo2 <> lp1 And Tmp_Echo3 <> lp1 And DBColumnList(lp1) = ComboBox23.Text Then
                                        Tmp_Echo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Tmp_Echo4 = lp1
                                        Exit For
                                    End If
                                Next
                            End If
                            If ComboBox24.Text <> "" Then
                                For lp1 = 0 To DataGridList.Length - 1
                                    If DataGridList(lp1) = "反響メモ" And Tmp_Echo1 <> lp1 And Tmp_Echo2 <> lp1 And Tmp_Echo3 <> lp1 And Tmp_Echo4 <> lp1 And DBColumnList(lp1) = ComboBox24.Text Then
                                        Tmp_Echo &= " " & Trim(MemberList(i)(lp1 + 1))
                                        Exit For
                                    End If
                                Next
                            End If

                            EchoUnion_flg = True
                        Else
                            If EchoUnion_flg = True And (DBColumnList(j) = ComboBox21.Text Or DBColumnList(j) = ComboBox22.Text Or DBColumnList(j) = ComboBox23.Text Or DBColumnList(j) = ComboBox24.Text) Then
                            Else
                                Tmp_Echo = Trim(MemberList(i)(j + 1))
                            End If
                        End If

                        '反響メモの場合、文字列100文字以内かチェック
                        If Trim(Tmp_Echo.Length) > 100 Then
                            'エラー内容を格納する。
                            'ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            'ErrorDataNo(ErrorDataNoCount) = j
                            'ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            'ErrorDataContent(ErrorDataNoCount) = "反響メモ"
                            'ErrorDataNoCount += 1
                            'ErrorData_Flg = True

                            '101文字以上なら、1～100文字目まで切り取る。
                            MemberList(i)(j + 1) = Tmp_Echo.Substring(0, 101)
                        Else
                            '作成するファイルにより格納値を変える。
                            If RadioButton1.Checked = True Then
                                '前後スペースを省いて格納する。
                                MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                            Else
                                If EchoUnion_flg = True Then
                                    MemberList(i)(Tmp_Echo1 + 1) = Tmp_Echo
                                Else
                                    MemberList(i)(j + 1) = Tmp_Echo
                                End If
                            End If

                        End If
                    ElseIf DataGridList(j) = "過去の来店回数" Then
                        '全角から半角に変換する。
                        MemberList(i)(j + 1) = StrConv(MemberList(i)(j + 1), VbStrConv.Narrow)

                        '過去の来店回数の場合、半角数値のみかチェック
                        If System.Text.RegularExpressions.Regex.IsMatch(Trim(MemberList(i)(j + 1)), "^[0-9]+$", _
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        Else
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "過去の来店回数"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        End If
                    ElseIf DataGridList(j) = "ランク" Then
                        'ランクの場合、文字列が20文字以内かチェックをする。
                        If Trim(MemberList(i)(j + 1).Length) > 20 Then
                            'エラー内容を格納する。
                            ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                            ErrorDataNo(ErrorDataNoCount) = j
                            ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                            ErrorDataContent(ErrorDataNoCount) = "ランク"
                            ErrorDataNoCount += 1
                            ErrorData_Flg = True
                            ErrorData_ImporChecktFLg = True
                        Else
                            '前後スペースを省いて格納する。
                            MemberList(i)(j + 1) = Trim(MemberList(i)(j + 1))
                        End If
                    End If
            Next
            '重複データかどうか。
            If MemberList(i)(j + 1) <> "" Then
                ReDim Preserve ErrorDataNo(0 To ErrorDataNoCount)
                ErrorDataNo(ErrorDataNoCount) = j
                ReDim Preserve ErrorDataContent(0 To ErrorDataNoCount)
                ErrorDataContent(ErrorDataNoCount) = "重複データNo"
                ErrorDataNoCount += 1
                ErrorData_Flg = True
                ErrorData_ImporChecktFLg = True
            End If

            If ErrorData_ImportFLg = True And ErrorData_ImporChecktFLg = False Then
                '空白にして取り込むデータを格納する。
                ReDim Preserve ErrorImportList(0 To ErrorImportDataCount)
                ReDim Preserve ErrorImportData(0 To ErrorImportMessageCount)
                k = 0
                'リストに出力するデータを格納
                For k = 0 To DataGridList.Length
                    ReDim Preserve ErrorImportDataList(0 To k)
                    For l = 0 To ErrorImportDataNo.Length - 1
                        'エラーのある項目は空白にする。
                        If k = ErrorImportDataNo(l) + 1 Then
                            ErrorImportDataList(k) = TmpErrorImportData(l)
                            Exit For
                        Else
                            ErrorImportDataList(k) = MemberList(i)(k)
                        End If
                    Next
                    'ErrorImportDataList(k) = MemberList(i)(k)
                Next

                ReDim Preserve ErrorImportDataList(0 To k)
                ErrorImportDataList(k) = MemberList(i)(k)
                ErrorImportData(ErrorImportMessageCount) = ErrorImportDataList

                For cnt = 0 To ErrorImportDataNo.Length - 1
                    'エラーのある項目名を格納(カンマ区切りで格納)
                    ReDim Preserve ErrorImportDataItem(0 To ErrorImportMessageCount)
                    If cnt = ErrorImportDataNo.Length - 1 Then
                        ErrorImportDataItem(ErrorImportMessageCount) &= ErrorImportDataNo(cnt)
                    Else
                        ErrorImportDataItem(ErrorImportMessageCount) &= ErrorImportDataNo(cnt) & ","
                    End If
                Next

                ReDim Preserve ErrorImportDataItemName(0 To ErrorImportMessageCount)
                For cnt = 0 To ErrorImportDataContent.Length - 1
                    'エラーのある項目名を格納(カンマ区切りで格納)
                    If cnt = ErrorImportDataContent.Length - 1 Then
                        ErrorImportDataItemName(ErrorImportMessageCount) &= ErrorImportDataContent(cnt)
                    Else
                        ErrorImportDataItemName(ErrorImportMessageCount) &= ErrorImportDataContent(cnt) & "、"
                    End If
                Next

                ErrorImportMessageCount += 1
                '不備、重複データを作成するためのフラグをTrueにする
                ExcelImportCreateFlg = True
            End If

            'ErrorData_FlgがTrueならエラーのあるデータの全項目を格納
            'If ErrorData_Flg = True Then
            'ANDだったのでOrに修正 2013/9/10 H.Mitao
            If ErrorData_Flg = True Or ErrorData_ImporChecktFLg = True Then
                ReDim Preserve ErrorData(0 To ErrorMessageCount)
                k = 0
                For k = 0 To DataGridList.Length
                    ReDim Preserve ErrorDataList(0 To k)
                    ErrorDataList(k) = MemberList(i)(k)
                Next
                ReDim Preserve ErrorDataList(0 To k)
                ErrorDataList(k) = MemberList(i)(k)
                ErrorData(ErrorMessageCount) = ErrorDataList

                For cnt = 0 To ErrorDataNo.Length - 1
                    'エラーのある項目名を格納(カンマ区切りで格納)
                    ReDim Preserve ErrorDataItem(0 To ErrorMessageCount)
                    If cnt = ErrorDataNo.Length - 1 Then
                        ErrorDataItem(ErrorMessageCount) &= ErrorDataNo(cnt)
                    Else
                        ErrorDataItem(ErrorMessageCount) &= ErrorDataNo(cnt) & ","
                    End If
                Next

                ReDim Preserve ErrorDataItemName(0 To ErrorMessageCount)
                For cnt = 0 To ErrorDataContent.Length - 1
                    'エラーのある項目名を格納(カンマ区切りで格納)
                    If cnt = ErrorDataContent.Length - 1 Then
                        ErrorDataItemName(ErrorMessageCount) &= ErrorDataContent(cnt)
                    Else
                        ErrorDataItemName(ErrorMessageCount) &= ErrorDataContent(cnt) & "、"
                    End If
                Next

                ErrorMessageCount += 1
                '不備、重複データを作成するためのフラグをTrueにする
                ExcelCreateFlg = True
                'Else

                'エラーがない、もしくは空白にすればエラーではなくなるデータか
            ElseIf ErrorData_Flg = False Or (ErrorData_Flg = True And RadioButton1.Checked = True) And (ErrorData_Flg = True And ErrorData_ImporChecktFLg = False And ErrorData_ImportFLg = True) Then
                'エラーがなかったら正常データとして配列に格納する
                'また、空白にして取り込むリストがあれば配列に格納する
                ReDim Preserve NormalDataList(0 To NormalDataCount)

                Dim DetailData() As String = Nothing
                ReDim DetailData(0 To 0)
                DetailData(0) = Trim(MemberList(i)(0))
                For lp = 1 To DataGridList.Length
                    ReDim Preserve DetailData(0 To lp)
                    If AdressUnion1_flg = True And (ComboBox8.Text = DBColumnList(lp - 1) Or ComboBox9.Text = DBColumnList(lp - 1) Or ComboBox10.Text = DBColumnList(lp - 1)) Then
                    ElseIf AdressUnion2_flg = True And (ComboBox12.Text = DBColumnList(lp - 1) Or ComboBox13.Text = DBColumnList(lp - 1) Or ComboBox14.Text = DBColumnList(lp - 1)) Then
                    Else
                        DetailData(lp) = MemberList(i)(lp)
                    End If
                Next
                ReDim Preserve DetailData(0 To lp)
                NormalDataList(NormalDataCount) = DetailData
                NormalDataCount += 1
            End If

        Next

        'Radiobutton1がtrueなら店舗向けデータの作成
        If RadioButton1.Checked = True Then
            dtNow = DateTime.Now
            'Trueならデータ不備、重複データリストの作成を行う
            If ExcelCreateFlg = True Then
                'エクセルファイルを作成する為の宣言
                Dim oExcel As Object
                oExcel = CreateObject("Excel.Application")
                Dim oBooks As Object
                oBooks = oExcel.WorkBooks
                Dim oBook As Object
                oBook = oBooks.Add
                Dim oSheets As Object
                oSheets = oBook.Worksheets
                Dim oSheet As Object
                oSheet = oSheets.Item(1)

                Dim PasteExcelData As Object = Nothing
                ReDim PasteExcelData(ErrorData.Length + 2, DBColumnList.Length + 3)

                'データを貼り付ける用変数
                Dim e_range As Excel.Range = Nothing
                Dim e_range1 As Excel.Range = Nothing
                Dim e_range2 As Excel.Range = Nothing

                Dim CreateFileName As String = Nothing

                '取り込んだデータファイル名取得
                FileName = TextBox1.Text
                FileName = FileName.Replace(".csv", "")

                'CSVファイル名と、項目行の設定
                CreateFileName = "\" & FileName & FileType1

                Try
                    PasteExcelData(0, 0) = "'エラー箇所"
                    PasteExcelData(0, 1) = "'PK"

                    'ヘッダー情報を1行目に書き込む。
                    For i = 0 To DBColumnList.Length - 1
                        PasteExcelData(0, i + 2) = "'" & DBColumnList(i)
                    Next
                    PasteExcelData(0, i + 2) = "'重複データNo"

                    '2行目以降はデータを書き込む
                    For x = 0 To ErrorData.Length - 1
                        PasteExcelData(x + 1, 0) = "'" & ErrorDataItemName(x)
                        'エラーのあった行のデータを書き込む。
                        y = 0
                        For y = 0 To DBColumnList.Length
                            PasteExcelData(x + 1, y + 1) = "'" & ErrorData(x)(y)
                        Next
                        '重複データを設定
                        PasteExcelData(x + 1, y + 1) = "'" & ErrorData(x)(y)

                        DataList = Split(ErrorDataItem(x), ",")

                        For z = 0 To UBound(DataList)
                            'エラー箇所の背景色を変更
                            oSheet.Cells(x + 2, DataList(z) + 3).Interior.ColorIndex = 7
                        Next
                    Next

                    '始点
                    e_range1 = DirectCast(oSheet.Cells(startY, startX), Excel.Range)
                    '終点
                    e_range2 = DirectCast(oSheet.Cells(startY + UBound(PasteExcelData, 1), startX + UBound(PasteExcelData, 2)), Excel.Range)
                    'セル範囲
                    e_range = oSheet.Range(e_range1, e_range2)
                    '貼り付け
                    e_range.Value = PasteExcelData

                    'ファイルを保存する
                    oBook.SaveAs(TextBox2.Text & CreateFileName)
                    oBook.Close(False)

                    '参照を解放(range,range1,range2)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range)
                    e_range = Nothing
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range1)
                    e_range1 = Nothing
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range2)
                    e_range2 = Nothing

                    '参照を解放(oBook)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook)
                    'oBookを解放
                    oBook = Nothing
                    '参照を解放(oBooks)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBooks)
                    oBooks = Nothing
                    '参照を解放(oSheet)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheet)
                    oSheet = Nothing
                    '参照を解放(oSheets)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheets)
                    oSheets = Nothing
                    '参照を解放(oExcel)
                    oExcel.Quit()
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel)
                    oExcel = Nothing
                Catch ex As Exception
                    '処理中ウインドウをDispose。
                    Processing.Dispose()
                    MsgBox("エクセルファイルの作成に失敗しました。")
                    oBook.Close(False)
                    '参照を解放(range,range1,range2)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range)
                    e_range = Nothing
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range1)
                    e_range1 = Nothing
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(e_range2)
                    e_range2 = Nothing

                    '参照を解放(oBook)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBook)
                    'oBookを解放
                    oBook = Nothing
                    '参照を解放(oBooks)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oBooks)
                    oBooks = Nothing
                    '参照を解放(oSheet)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheet)
                    oSheet = Nothing
                    '参照を解放(oSheets)
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oSheets)
                    oSheets = Nothing
                    '参照を解放(oExcel)
                    oExcel.Quit()
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oExcel)
                    oExcel = Nothing
                    Exit Sub
                End Try

                '処理中ウインドウをDispose。
                Processing.Dispose()
                MsgBox("データに不備があった為、ファイルを作成しました。" & vbCr & "作成ファイル名：" & FileName & FileType1)
            End If
            '処理中ウインドウを表示
            Processing.Show()
            'Processingウインドウの「処理中です。」のメッセージをリフレッシュさせることにより表示
            Processing.Label1.Refresh()

            If NormalDataCount = 0 Then
                MsgBox("取り込みデータが0件だった為、" & vbCr & "取り込み用データファイルの作成は行われませんでした。")

                '処理中ウインドウをDispose。
                Processing.Dispose()
                '指定のフォルダを開く
                'System.Diagnostics.Process.Start(TextBox2.Text)
                Exit Sub
            End If

            '取り込みを行うデータ（正常なデータ）の作成を行う
            'エクセルファイルを作成する為の宣言
            Dim oExcel2 As Object
            oExcel2 = CreateObject("Excel.Application")
            Dim oBooks2 As Object
            oBooks2 = oExcel2.WorkBooks
            Dim oBook2 As Object
            oBook2 = oBooks2.Add
            Dim oSheets2 As Object
            oSheets2 = oBook2.Worksheets
            Dim oSheet2 As Object
            oSheet2 = oSheets2.Item(1)

            'データを貼り付ける用変数
            Dim range As Excel.Range = Nothing
            Dim range1 As Excel.Range = Nothing
            Dim range2 As Excel.Range = Nothing

            Dim PasteData As Object = Nothing
            ReDim PasteData(NormalDataList.Length + 1, DBColumnList.Length + 1)

            Dim CreateFileName2 As String = Nothing

            '取り込んだデータファイル名取得
            FileName = TextBox1.Text
            FileName = FileName.Replace(".csv", "")
            CreateFileName2 = "\" & FileName & FileType2

            '処理開始時刻を記憶
            'Dim start As DateTime = Now
            Try
                'ヘッダー情報を1行目に書き込む。
                PasteData(0, 0) = "'PK"
                '項目数分ループ
                For i = 0 To DBColumnList.Length - 1
                    PasteData(0, i + 1) = "'" & DBColumnList(i)
                Next
                'データ件数分ループ
                For i = 0 To NormalDataList.length - 1
                    'データを書き込む。
                    For j = 0 To DBColumnList.Length
                        PasteData(i + 1, j) = "'" & NormalDataList(i)(j)
                    Next
                Next

                '始点
                range1 = DirectCast(oSheet2.Cells(startY, startX), Excel.Range)
                '終点
                range2 = DirectCast(oSheet2.Cells(startY + UBound(PasteData, 1), startX + UBound(PasteData, 2)), Excel.Range)
                'セル範囲
                range = oSheet2.Range(range1, range2)
                '貼り付け
                range.Value = PasteData

                '終了時間
                'Dim span As TimeSpan = Now - start
                'MessageBox.Show(String.Format("処理時間は{0}秒です。", span.TotalSeconds))

                'ファイルを保存し、Closeする。
                oBook2.SaveAs(TextBox2.Text & CreateFileName2)
                oBook2.Close(False)
                '参照の解放
                Excel_Release(range, range1, range2, oBook2, oBooks2, oSheet2, oSheets2, oExcel2)

            Catch ex As Exception
                '処理中ウインドウをDispose。

                Processing.Dispose()
                MsgBox("エクセルファイルの作成に失敗しました。")

                '参照の解放
                Excel_Release(range, range1, range2, oBook2, oBooks2, oSheet2, oSheets2, oExcel2)

                Exit Sub
            End Try

            '処理中ウインドウをDispose。
            Processing.Dispose()

            MsgBox("取り込みデータを作成しました。" & vbCr & "作成ファイル名：" & FileName & FileType2)
            '指定のフォルダを開く
            'System.Diagnostics.Process.Start(TextBox2.Text)

        Else
            'Radiobutton2がTrueなら五十嵐さんに提出用のDD向けフォーマットファイルの作成

            '処理中ウインドウを表示
            Processing.Show()
            'Processingウインドウの「処理中です。」のメッセージをリフレッシュさせることにより表示
            Processing.Label1.Refresh()

            '作成するファイル名を格納
            Dim CreateFileName As String = Nothing
            '取り込んだデータファイル名取得
            FileName = TextBox1.Text
            '拡張子を取り除く
            FileName = FileName.Replace(".csv", "")
            'xlsファイル名と、項目行の設定
            CreateFileName = "\" & FileName & FileType3

            If NormalDataCount = 0 Then
                MsgBox("作成するデータがない為、処理を終了します。")
                Exit Sub
            End If

            '処理開始時刻を記憶
            'Dim start As DateTime = Now

            Dim tmpNormalDataList As Object = Nothing
            ReDim tmpNormalDataList(NormalDataList.Length + 1, FormatList.Length - 1)

            '項目数分ループし１行目に項目を設定
            For i = 0 To FormatList.Length - 1
                tmpNormalDataList(0, i) = FormatList(i)
            Next
            'データ件数分ループさせ、項目名とデータが一致したら、エクセルに書き込む
            For i = 0 To NormalDataList.length - 1
                For j = 0 To DataGridList.Length - 1
                    For k = 0 To FormatList.Length - 1
                        'DataGridViewのチェック項目とデータコンバートのフィールド名が一致したら書き込み
                        If DataGridList(j) = FormatList(k) Then
                            '項目を設定
                            '姓名を分割するを選択し、顧客名（姓）の項目の場合は分割を行う。
                            If ComboBox6.Text = "分割する" And (FormatList(k) = "顧客名(姓)" Or FormatList(k) = "ふりがな(姓)") Then
                                TmpData = Nothing
                                NormalDataList(i)(j + 1) = NormalDataList(i)(j + 1).Replace("　", " ")
                                TmpData = NormalDataList(i)(j + 1).split(" ")
                                ReDim Preserve TmpData(0 To 1)
                                tmpNormalDataList(i + 1, k) = "'" & Trim(TmpData(0))
                                If (IsDBNull(TmpData(1))) Then
                                Else
                                    tmpNormalDataList(i + 1, k + 1) = "'" & Trim(TmpData(1))
                                End If
                                Exit For
                            Else
                                If AdressUnion1_flg = True And (ComboBox8.Text = DBColumnList(j) Or ComboBox9.Text = DBColumnList(j) Or ComboBox10.Text = DBColumnList(j)) Then
                                ElseIf AdressUnion2_flg = True And (ComboBox12.Text = DBColumnList(j) Or ComboBox13.Text = DBColumnList(j) Or ComboBox14.Text = DBColumnList(j)) Then
                                ElseIf MemoUnion_flg = True And (ComboBox16.Text = DBColumnList(j) Or ComboBox17.Text = DBColumnList(j) Or ComboBox18.Text = DBColumnList(j) Or ComboBox19.Text = DBColumnList(j)) Then
                                ElseIf EchoUnion_flg = True And (ComboBox21.Text = DBColumnList(j) Or ComboBox22.Text = DBColumnList(j) Or ComboBox23.Text = DBColumnList(j) Or ComboBox24.Text = DBColumnList(j)) Then
                                Else
                                    tmpNormalDataList(i + 1, k) = "'" & NormalDataList(i)(j + 1)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Next
            Next

            'エクセルファイルを作成する為の宣言
            Dim oExcel As Object
            oExcel = CreateObject("Excel.Application")
            Dim oBooks As Object
            oBooks = oExcel.WorkBooks
            Dim oBook As Object
            oBook = oBooks.Add
            Dim oSheets As Object
            oSheets = oBook.Worksheets
            Dim oSheet As Object
            oSheet = oSheets.Item(1)

            'データを貼り付ける用変数
            Dim range As Excel.Range = Nothing
            Dim range1 As Excel.Range = Nothing
            Dim range2 As Excel.Range = Nothing

            Try
                '始点
                range1 = DirectCast(oSheet.Cells(startY, startX), Excel.Range)
                '終点
                range2 = DirectCast(oSheet.Cells(startY + UBound(tmpNormalDataList, 1), startX + UBound(tmpNormalDataList, 2)), Excel.Range)
                'セル範囲
                range = oSheet.Range(range1, range2)
                '貼り付け
                range.Value = tmpNormalDataList

                '処理時間を計算（現在時刻－処理開始時刻）
                'Dim span As TimeSpan = Now - start
                'MessageBox.Show(String.Format("処理時間は{0}秒です。", span.TotalSeconds))

                'ファイルを保存し、Closeする。
                oBook.SaveAs(TextBox2.Text & CreateFileName)
                'oBookをcloseする。
                oBook.Close(False)

                '参照の解放
                Excel_Release(range, range1, range2, oBook, oBooks, oSheet, oSheets, oExcel)

            Catch ex As Exception
                '処理中ウインドウをDispose。
                Processing.Dispose()
                MsgBox("エクセルファイルの作成に失敗しました。")

                'ファイルを保存し、Closeする。
                oBook.SaveAs(TextBox2.Text & CreateFileName)
                'oBookをcloseする。
                oBook.Close(False)

                '参照の解放
                Excel_Release(range, range1, range2, oBook, oBooks, oSheet, oSheets, oExcel)
                Exit Sub
            End Try

            '処理中ウインドウを表示
            Processing.Show()
            'Processingウインドウの「処理中です。」のメッセージをリフレッシュさせることにより表示
            Processing.Label1.Refresh()

            MsgBox("DD取り込みデータを作成しました。" & vbCr & "作成ファイル名：" & FileName & FileType3)

            '空白にして取り込むリストがあれば
            If ExcelImportCreateFlg = True Then

                '取り込みを行うデータ（正常なデータ）の作成を行う
                'エクセルファイルを作成する為の宣言
                Dim oExcel2 As Object
                oExcel2 = CreateObject("Excel.Application")
                Dim oBooks2 As Object
                oBooks2 = oExcel2.WorkBooks
                Dim oBook2 As Object
                oBook2 = oBooks2.Add
                Dim oSheets2 As Object
                oSheets2 = oBook2.Worksheets
                Dim oSheet2 As Object
                oSheet2 = oSheets2.Item(1)

                Dim PasteExcelData As Object = Nothing
                ReDim PasteExcelData(ErrorImportData.Length + 2, DBColumnList.Length + 3)

                'データを貼り付ける用変数
                Dim e_range As Excel.Range = Nothing
                Dim e_range1 As Excel.Range = Nothing
                Dim e_range2 As Excel.Range = Nothing

                Dim e_CreateFileName As String = Nothing

                '取り込んだデータファイル名取得
                FileName = TextBox1.Text
                FileName = FileName.Replace(".csv", "")

                'CSVファイル名と、項目行の設定
                CreateFileName = "\" & FileName & FileType4

                Try
                    PasteExcelData(0, 0) = "'エラー箇所"
                    PasteExcelData(0, 1) = "'PK"

                    'ヘッダー情報を1行目に書き込む。
                    For i = 0 To DBColumnList.Length - 1
                        PasteExcelData(0, i + 2) = "'" & DBColumnList(i)
                    Next

                    '2行目以降はデータを書き込む
                    For x = 0 To ErrorImportData.Length - 1
                        PasteExcelData(x + 1, 0) = "'" & ErrorImportDataItemName(x)
                        'エラーのあった行のデータを書き込む。
                        y = 0
                        For y = 0 To DBColumnList.Length
                            PasteExcelData(x + 1, y + 1) = "'" & ErrorImportData(x)(y)
                        Next

                        DataList = Split(ErrorImportDataItem(x), ",")

                        For z = 0 To UBound(DataList)
                            'エラー箇所の背景色を変更
                            oSheet2.Cells(x + 2, DataList(z) + 3).Interior.ColorIndex = 7
                        Next
                    Next

                    '始点
                    e_range1 = DirectCast(oSheet2.Cells(startY, startX), Excel.Range)
                    '終点
                    e_range2 = DirectCast(oSheet2.Cells(startY + UBound(PasteExcelData, 1), startX + UBound(PasteExcelData, 2)), Excel.Range)
                    'セル範囲
                    e_range = oSheet2.Range(e_range1, e_range2)
                    '貼り付け
                    e_range.Value = PasteExcelData

                    'ファイルを保存する
                    oBook2.SaveAs(TextBox2.Text & CreateFileName)
                    oBook2.Close(False)

                    '参照の解放
                    Excel_Release(range, range1, range2, oBook2, oBooks2, oSheet2, oSheets2, oExcel2)

                Catch ex As Exception
                    '処理中ウインドウをDispose。
                    Processing.Dispose()
                    MsgBox("エクセルファイルの作成に失敗しました。")
                    oBook2.Close(False)

                    '参照の解放
                    Excel_Release(e_range, e_range1, e_range2, oBook2, oBooks2, oSheet2, oSheets2, oExcel2)
                    Exit Sub
                End Try

                MsgBox("DD追加取り込み報告データを作成しました。" & vbCr & "作成ファイル名：" & FileName & FileType4)
                '指定のフォルダを開く
                'System.Diagnostics.Process.Start(TextBox2.Text)
            End If

            '処理中ウインドウをDispose。
            Processing.Dispose()

        End If

        '指定のフォルダを開く
        'System.Diagnostics.Process.Start(TextBox2.Text)

    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'ダイアログ設定
        Dim closecheck As DialogResult = MessageBox.Show("システムを終了してもよろしいですか？", _
                                                     "確認", _
                                                     MessageBoxButtons.YesNo, _
                                                     MessageBoxIcon.Question)
        If closecheck = DialogResult.Yes Then
            'ダイアログで「はい」が選択された時 
            Application.Exit()
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'アプリケーションの終了
        Application.Exit()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '参照ボタンが押された時の処理

        'Dim fbd As New FolderBrowserDialog
        ''上部に表示する説明テキストを指定する
        'fbd.Description = "フォルダを指定してください。"
        ''ルートフォルダを指定する
        ''デフォルトでDesktop
        'fbd.RootFolder = Environment.SpecialFolder.Desktop
        ''最初に選択するフォルダを指定する
        ''RootFolder以下にあるフォルダである必要がある
        'fbd.SelectedPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal)
        ''ユーザーが新しいフォルダを作成できるようにする
        ''デフォルトでTrue
        'fbd.ShowNewFolderButton = True

        ''ダイアログを表示する
        'If fbd.ShowDialog(Me) = DialogResult.OK Then
        '    '選択したpathを表示
        '    TextBox2.Text = fbd.SelectedPath
        'End If

        If Trim(TextBox2.Text) <> "" Then
            System.Diagnostics.Process.Start(TextBox2.Text)
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        '不備・重複、取り込み可能リストの作成（店舗へのデータ確認用）ボタンが押されたら、
        '五十嵐さん提出用フォーマットにてリストを作成（DD取り込み用）ボタンをFalseにする。
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        '五十嵐さん提出用フォーマットにてリストを作成（DD取り込み用）ボタンが押されたら、
        '不備・重複、取り込み可能リストの作成（店舗へのデータ確認用）ボタンをFalseにする。
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        'AND検索ボタンが押されたら、OR検索ボタンをFalseにする。
        If RadioButton3.Checked = True Then
            RadioButton4.Checked = False
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        'OR検索ボタンが押されたら、AND検索ボタンをFalseにする。
        If RadioButton4.Checked = True Then
            RadioButton3.Checked = False
        End If
    End Sub
End Class