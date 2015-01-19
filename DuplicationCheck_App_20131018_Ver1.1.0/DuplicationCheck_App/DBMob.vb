Imports System.Data.OleDb

Module DBMob
    '項目名を格納
    Public SQLColumn As String = Nothing
    'データを格納
    Public SQLData() As String = Nothing
    'カラム数を格納
    Public ColumnCount As Integer
    'C_の項目も含めたカラム数
    Public ALL_ColumnCount As Integer
    '項目名リスト
    Public ColumnList() As String
    'コンボボックスのプルダウン情報を格納
    Public DBColumnList() As String

    Public Structure CSV_List
        Dim PK As Integer
        '読み込んだデータをそのまま格納するフィールド
        Dim MemberID As String
        Dim StoreName As String
        Dim memberID2 As String
        Dim Name As String
        Dim Abbreviation As String
        Dim NameKana As String
        Dim MemberNo As String
        Dim TEL As String
        Dim RegistDay As String
        Dim Post As String
        Dim Address1 As String
        Dim Address2 As String
        Dim Address3 As String
        Dim Address4 As String
        Dim Birth As String
        Dim Sex As String
        Dim FirstCome As String
        Dim Updateday As String
        Dim Item1 As String
        Dim Item2 As String
        Dim Item3 As String
        Dim StoreCode As String
        Dim Rank As String
        '読み込んだデータに対し、全角->半角、スペース削除をしたものを格納するフィールド
        Dim C_MemberID As String
        Dim C_StoreName As String
        Dim C_memberID2 As String
        Dim C_Name As String
        Dim C_Abbreviation As String
        Dim C_NameKana As String
        Dim C_MemberNo As String
        Dim C_TEL As String
        Dim C_RegistDay As String
        Dim C_Post As String
        Dim C_Address1 As String
        Dim C_Address2 As String
        Dim C_Address3 As String
        Dim C_Address4 As String
        Dim C_Birth As String
        Dim C_Sex As String
        Dim C_FirstCome As String
        Dim C_Updateday As String
        Dim C_Item1 As String
        Dim C_Item2 As String
        Dim C_Item3 As String
        Dim C_StoreCode As String
        Dim C_Rank As String
    End Structure
    '重複チェックを行う項目を格納
    Public Structure Duplicate_List
        Dim Check1 As String
        Dim Check2 As String
        Dim Check3 As String
        Dim Check4 As String
        Dim Check5 As String
    End Structure
    'mdbファイルを作成する
    Function MDB_CRTDATABASE(ByVal PistrMakPath As String, _
                             ByVal DBColumnList() As String) As Boolean
        Dim con As New OleDbConnection()
        Dim cmd As New OleDbCommand()
        'カタログ
        Dim objCat As ADOX.Catalog
        'データベースパラメータ
        Dim strDatbasePara As String
        'テーブル名称
        Dim strTable As String
        'テーブル
        Dim objTable As ADOX.Table
        Dim i As Integer = 0

        MDB_CRTDATABASE = False

        strDatbasePara = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
                        "Data Source=" + PistrMakPath + ";" + _
                        "Jet OLEDB:Engine Type=5;"
        Try
            'ADOXオブジェクトを作成します
            objCat = New ADOX.Catalog
            'MDB作成
            objCat.Create(strDatbasePara)
            '項目用初期テーブルを作成します
            ' テーブル名を指定してテーブルを追加する
            strTable = "MemberTable"
            objTable = New ADOX.Table

            With objTable
                .Name = strTable
                .Columns.Append("PK", ADOX.DataTypeEnum.adInteger)
                For i = 0 To DBColumnList.Length - 1
                    'メモ型
                    .Columns.Append(DBColumnList(i), ADOX.DataTypeEnum.adLongVarWChar)
                    'NULL許可設定
                    .Columns(DBColumnList(i)).Attributes = ADOX.ColumnAttributesEnum.adColNullable
                Next
                '重複データNo、テキスト型　設定
                .Columns.Append("重複データNo", ADOX.DataTypeEnum.adLongVarWChar)
                '重複データNo NULL許可
                .Columns("重複データNo").Attributes = ADOX.ColumnAttributesEnum.adColNullable
            End With
            objCat.Tables.Append(objTable)
            SQLColumn = ""
            For i = 0 To DBColumnList.Length - 1
                If i = DBColumnList.Length - 1 Then
                    SQLColumn &= "[" & DBColumnList(i) & "]"
                Else
                    SQLColumn &= "[" & DBColumnList(i) & "],"
                End If
            Next
            MDB_CRTDATABASE = True
        Catch ex As Exception
        End Try
        objCat = Nothing
    End Function
    'MDBにデータを登録する
    Function MDB_INSERT(ByVal LoopControl As Integer, _
                        ByVal PistrMakPath As String, _
                        ByVal SQLColumn As String, _
                        ByVal SQLData() As String) As Boolean
        Dim con As New OleDbConnection()
        Dim cmd As New OleDbCommand()
        Dim result As Boolean = True
        'データベースパラメータ
        Dim strDatbasePara As String
        Dim trans As OleDb.OleDbTransaction = Nothing
        Dim Count As Long = 0

        MDB_INSERT = True

        strDatbasePara = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
                "Data Source=" + PistrMakPath + ";" + _
                "Jet OLEDB:Engine Type=5;"

        con.ConnectionString = strDatbasePara

        Try
            ' コネクションの設定
            cmd.Connection = con
            ' DB接続を開く
            con.Open()
            'begin
            trans = con.BeginTransaction()
            'コマンドオブジェクトにトランザクション関連付け
            cmd.Transaction = trans
            ' SQL文の設定
            For Count = 0 To SQLData.Length - 1
                cmd.CommandText = "Insert into MemberTable([PK],"
                cmd.CommandText &= SQLColumn
                cmd.CommandText &= ",[重複データNo])values("
                cmd.CommandText &= Count + LoopControl
                cmd.CommandText &= ","
                cmd.CommandText &= SQLData(Count)
                cmd.CommandText &= ",'');"
                ' Insert実行 
                result = cmd.ExecuteNonQuery()
            Next
            'コミット
            trans.Commit()
            MDB_INSERT = True
        Catch ex As Exception
            MDB_INSERT = False
            trans.Rollback()
            MsgBox(ex.Message)
            Exit Function
        Finally
            con.Close()
            con.Dispose()
        End Try
    End Function
    'メンバーデータの取得
    Function MDB_GET_MemberData(ByVal PistrMakPath As String, _
                                ByRef MemberList() As Object, _
                                ByRef Result As Boolean, _
                                ByRef ErrorMessage As String) As Boolean

        Dim con As New OleDbConnection()
        Dim cmd As New OleDbCommand()
        Dim DuplicateData() As Duplicate_List = Nothing
        Dim DuplicateData_Modify() As Duplicate_List = Nothing
        'データベースパラメータ
        Dim strDatbasePara As String
        Dim trans As OleDb.OleDbTransaction = Nothing
        '取得データ格納
        Dim DataList As OleDbDataReader
        '配列の再設定用のデータ数格納
        Dim DataCount As Integer = 0
        'ループ用変数
        Dim lp As Integer = 0

        strDatbasePara = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + PistrMakPath + ";" + "Jet OLEDB:Engine Type=5;"

        con.ConnectionString = strDatbasePara

        MDB_GET_MemberData = True

        Dim TmpDataCount As Integer = 50000
        'ループの旅にRedim Preserveすると処理が遅くなるので、事前に一時的に配列を作成。
        ReDim MemberList(0 To TmpDataCount)
        ' コネクションの設定
        cmd.Connection = con
        ' DB接続を開く
        con.Open()

        Try
            ' SQL文の設定
            cmd.CommandText = "SELECT PK,"
            cmd.CommandText &= SQLColumn
            cmd.CommandText &= ",重複データNo From MemberTable ORDER BY PK"

            DataList = cmd.ExecuteReader()
            'データ件数分ループ
            While DataList.Read()
                '配列の再設定（ループのたびに作成しなおすと遅くなるので5万件単位）
                If TmpDataCount < DataCount Then
                    TmpDataCount = TmpDataCount + 50000
                    ReDim Preserve MemberList(0 To TmpDataCount)
                End If

                Dim DetailData() As String = Nothing
                ReDim DetailData(0 To 0)
                DetailData(0) = DataList("PK")
                For lp = 1 To ALL_ColumnCount
                    ReDim Preserve DetailData(0 To lp)
                    If IsDBNull(DataList(DBColumnList(lp - 1))) Then
                        DetailData(lp) = ""
                    Else
                        DetailData(lp) = DataList(DBColumnList(lp - 1))
                    End If
                Next
                '重複データNo用に配列の再設定
                ReDim Preserve DetailData(0 To lp)
                If IsDBNull(DataList("重複データNo")) Then
                    DetailData(lp) = ""
                Else
                    DetailData(lp) = DataList("重複データNo")
                End If
                'データ格納
                MemberList(DataCount) = DetailData

                DataCount = DataCount + 1
            End While
            DataList.Close()

            'ループ終了後に、配列を半端分をなくす。
            ReDim Preserve MemberList(0 To DataCount - 1)
        Catch ex As Exception
            MDB_GET_MemberData = False
            MsgBox(ex.Message)
            Exit Function
        Finally
            con.Close()
            con.Dispose()
        End Try
    End Function
    '重複チェック
    Function MDB_DuplicateCheck(ByVal PistrMakPath As String, _
                                ByVal Check1 As String, _
                                ByVal Check2 As String, _
                                ByVal Check3 As String, _
                                ByVal Check4 As String, _
                                ByVal Check5 As String, _
                                ByVal Search_Type As String, _
                                ByRef Result As Boolean, _
                                ByRef ErrorMessage As String) As Boolean

        Dim con As New OleDbConnection()
        Dim cmd As New OleDbCommand()
        Dim DuplicateData() As Duplicate_List = Nothing
        Dim DuplicateData_Modify() As Duplicate_List = Nothing
        'データベースパラメータ
        Dim strDatbasePara As String

        Dim trans As OleDb.OleDbTransaction = Nothing
        'データ格納用
        Dim DataList As OleDbDataReader
        '重複データ格納用
        Dim DuplicateDataList As OleDbDataReader

        Dim DataCount As Integer = 0

        'SELECTの項目を格納
        Dim SQLGroup As String = Nothing
        'HAVINEの内容を格納
        Dim SQLHAVINE As String = Nothing

        'ループ用
        Dim i As Integer = 0
        Dim x As Integer = 0
        '重複データのPKを格納
        Dim DuplicatePK As String = Nothing
        'PKのリストを格納
        Dim PK_List() As Integer = Nothing

        '重複チェック項目の格納
        Dim check() As String = Nothing

        strDatbasePara = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + PistrMakPath + ";" + "Jet OLEDB:Engine Type=5;"

        con.ConnectionString = strDatbasePara

        MDB_DuplicateCheck = True

        Try
            ' コネクションの設定
            cmd.Connection = con
            ' DB接続を開く
            con.Open()

            trans = con.BeginTransaction()

            'コマンドオブジェクトにトランザクション関連付け
            cmd.Transaction = trans

            If Search_Type = "OR検索" Then
                ReDim check(0 To 4)
                check(0) = Check1
                check(1) = Check2
                check(2) = Check3
                check(3) = Check4
                check(4) = Check5

                For i = 0 To 4
                    If check(i) <> "" Then

                        SQLGroup = check(i)

                        ' SQL文の設定
                        cmd.CommandText = "SELECT "
                        cmd.CommandText &= SQLGroup
                        cmd.CommandText &= " FROM MemberTable GROUP BY "
                        cmd.CommandText &= SQLGroup
                        cmd.CommandText &= " HAVING COUNT("
                        cmd.CommandText &= SQLGroup
                        cmd.CommandText &= ") > 1 "
                        'SQL実行
                        DataList = cmd.ExecuteReader()

                        x = 0
                        'データ件数分ループ
                        While DataList.Read()
                            ReDim Preserve DuplicateData(0 To x)
                            DuplicateData(x).Check1 = DataList(SQLGroup).Replace("'", "''")
                            x += 1
                        End While
                        '破棄
                        DataList.Close()

                        For y = 0 To DuplicateData.Length - 1

                            '該当した重複しているPKを取り出す
                            cmd.CommandText = "SELECT PK FROM MemberTable WHERE "
                            cmd.CommandText &= SQLGroup
                            cmd.CommandText &= "='"
                            cmd.CommandText &= DuplicateData(y).Check1
                            cmd.CommandText &= "'"

                            DuplicateDataList = cmd.ExecuteReader()

                            DataCount = 0
                            ReDim PK_List(0 To DataCount)
                            While DuplicateDataList.Read()
                                ReDim Preserve PK_List(0 To DataCount)
                                PK_List(DataCount) = DuplicateDataList("PK")
                                DataCount = DataCount + 1
                            End While

                            DuplicateDataList.Close()

                            DuplicatePK = ""

                            '重複したPKを：区切りでつなげる。
                            For j = 0 To PK_List.Length - 1
                                If j = PK_List.Length - 1 Then
                                    DuplicatePK &= PK_List(j)
                                Else
                                    DuplicatePK &= PK_List(j) & ":"
                                End If
                            Next

                            '重複№情報をUPDATEする。
                            For k = 0 To PK_List.Length - 1
                                cmd.CommandText = "UPDATE MemberTable SET 重複データNo ='"
                                cmd.CommandText &= DuplicatePK
                                cmd.CommandText &= "' WHERE PK="
                                cmd.CommandText &= PK_List(k)
                                cmd.CommandText &= ";"

                                ' UPDATE実行 
                                Result = cmd.ExecuteNonQuery()
                            Next
                        Next
                    End If
                Next
            ElseIf Search_Type = "AND検索" Then
                SQLGroup = Check1
                SQLHAVINE = Check1

                If Check2 <> "" Then
                    SQLGroup &= "," & Check2
                    SQLHAVINE &= " AND " & Check2
                End If
                If Check3 <> "" Then
                    SQLGroup &= "," & Check3
                    SQLHAVINE &= " AND " & Check3
                End If
                If Check4 <> "" Then
                    SQLGroup &= "," & Check4
                    SQLHAVINE &= " AND " & Check4
                End If
                If Check5 <> "" Then
                    SQLGroup &= "," & Check5
                    SQLHAVINE &= " AND " & Check5
                End If

                'SQL文の設定
                cmd.CommandText = "SELECT "
                cmd.CommandText &= SQLGroup
                cmd.CommandText &= " FROM MemberTable GROUP BY "
                cmd.CommandText &= SQLGroup
                cmd.CommandText &= " HAVING COUNT("
                cmd.CommandText &= SQLHAVINE
                cmd.CommandText &= ") > 1 "
                DataList = cmd.ExecuteReader()

                'ループ内でRedim Preserveを行うと遅くなるので一旦5000で作成
                Dim MaxArray As Integer = 5000
                ReDim Preserve DuplicateData(0 To MaxArray)

                While DataList.Read()
                    If i > MaxArray Then
                        MaxArray = MaxArray + 5000
                        ReDim Preserve DuplicateData(0 To MaxArray)
                    End If

                    DuplicateData(i).Check1 = DataList(Check1).Replace("'", "''")
                    If Check2 <> "" Then
                        DuplicateData(i).Check2 = DataList(Check2).Replace("'", "''")
                    End If
                    If Check3 <> "" Then
                        DuplicateData(i).Check3 = DataList(Check3).Replace("'", "''")
                    End If
                    If Check4 <> "" Then
                        DuplicateData(i).Check4 = DataList(Check4).Replace("'", "''")
                    End If
                    If Check5 <> "" Then
                        DuplicateData(i).Check5 = DataList(Check5).Replace("'", "''")
                    End If

                    i = i + 1
                End While

                Dim MaxArrayCount = i - 1
                '最終的な配列数を確定させる。
                ReDim Preserve DuplicateData(0 To MaxArrayCount)

                'Close
                DataList.Close()

                If i > 0 Then
                    For j = 0 To DuplicateData.Length - 1
                        '該当した重複している名前のPKを取り出す為のSQL作成
                        cmd.CommandText = "SELECT PK FROM MemberTable WHERE "
                        cmd.CommandText &= Check1
                        cmd.CommandText &= "='"
                        cmd.CommandText &= DuplicateData(j).Check1
                        cmd.CommandText &= "'"
                        If Check2 <> "" Then
                            cmd.CommandText &= " AND "
                            cmd.CommandText &= Check2
                            cmd.CommandText &= "='"
                            cmd.CommandText &= DuplicateData(j).Check2
                            cmd.CommandText &= "'"
                        End If
                        If Check3 <> "" Then
                            cmd.CommandText &= " AND "
                            cmd.CommandText &= Check3
                            cmd.CommandText &= "='"
                            cmd.CommandText &= DuplicateData(j).Check3
                            cmd.CommandText &= "'"
                        End If
                        If Check4 <> "" Then
                            cmd.CommandText &= " AND "
                            cmd.CommandText &= Check4
                            cmd.CommandText &= "='"
                            cmd.CommandText &= DuplicateData(j).Check4
                            cmd.CommandText &= "'"
                        End If
                        If Check5 <> "" Then
                            cmd.CommandText &= " AND "
                            cmd.CommandText &= Check5
                            cmd.CommandText &= "='"
                            cmd.CommandText &= DuplicateData(j).Check5
                            cmd.CommandText &= "'"
                        End If
                        DataCount = 0

                        DuplicateDataList = cmd.ExecuteReader()

                        Dim MaxArrayDataListCount = 5000

                        ReDim Preserve PK_List(0 To MaxArrayDataListCount)

                        While DuplicateDataList.Read()
                            If DataCount > MaxArrayDataListCount Then
                                MaxArrayDataListCount = MaxArrayDataListCount + 5000
                                'ReDim Preserve DuplicateData(0 To MaxArrayDataListCount)
                                ReDim Preserve PK_List(0 To MaxArrayDataListCount)
                            End If

                            PK_List(DataCount) = DuplicateDataList("PK")
                            DataCount = DataCount + 1
                        End While

                        MaxArrayCount = DataCount - 1
                        '最終的な配列数を確定させる。
                        ReDim Preserve PK_List(0 To DataCount - 1)

                        DuplicateDataList.Close()

                        DuplicatePK = ""

                        '重複したPKを：区切りでつなげる。
                        For i = 0 To PK_List.Length - 1
                            If i = PK_List.Length - 1 Then
                                DuplicatePK &= PK_List(i)
                            Else
                                DuplicatePK &= PK_List(i) & ":"
                            End If
                        Next

                        '重複№情報をUPDATEする。
                        For i = 0 To PK_List.Length - 1
                            cmd.CommandText = "UPDATE MemberTable SET 重複データNo ='"
                            cmd.CommandText &= DuplicatePK
                            cmd.CommandText &= "' WHERE PK="
                            cmd.CommandText &= PK_List(i)
                            cmd.CommandText &= ";"
                            ' UPDATE実行 
                            Result = cmd.ExecuteNonQuery()
                        Next
                    Next
                End If
            End If
            'commit
            trans.Commit()
        Catch ex As Exception
            MDB_DuplicateCheck = False
            MsgBox(ex.Message)
            Exit Function
        Finally
            con.Close()
            con.Dispose()
        End Try
    End Function

    '重複Noのクリア
    Function MDB_DuplicateClear(ByVal PistrMakPath As String, _
                                ByRef Result As Boolean, _
                                ByRef ErrorMessage As String) As Boolean

        Dim con As New OleDbConnection()
        Dim cmd As New OleDbCommand()
        Dim DuplicateData() As Duplicate_List = Nothing
        Dim DuplicateData_Modify() As Duplicate_List = Nothing
        'データベースパラメータ
        Dim strDatbasePara As String

        Dim trans As OleDb.OleDbTransaction = Nothing

        strDatbasePara = "Provider=Microsoft.Jet.OLEDB.4.0;" + _
        "Data Source=" + PistrMakPath + ";" + _
        "Jet OLEDB:Engine Type=5;"

        con.ConnectionString = strDatbasePara

        MDB_DuplicateClear = True

        Try
            ' コネクションの設定
            cmd.Connection = con
            ' DB接続を開く
            con.Open()

            trans = con.BeginTransaction()

            'コマンドオブジェクトにトランザクション関連付け
            cmd.Transaction = trans

            cmd.CommandText = "UPDATE MemberTable SET 重複データNo ='';"
            ' UPDATE実行 
            Result = cmd.ExecuteNonQuery()
            'commit
            trans.Commit()
        Catch ex As Exception
            MDB_DuplicateClear = False
            MsgBox(ex.Message)
            Exit Function
        Finally
            con.Close()
            con.Dispose()
        End Try
    End Function
End Module
