'-------------------------------------------------------------------------------
' Random Learning
'
' 概要
' ・英単語などを暗記するためのツールです。
'
' 使い方
' ・Question列に質問、Answer列に解答を入力した上で、[エクササイズ開始]ボタンクリックします (Main関数がスタートします)
' ・Questionがメッセージボックスで表示されたら、そのAnswerを考えた後、[OK]ボタンを押してください。
' ・Answerがメッセージボックスで表示されたら、正解の場合は[OK]ボタン、不正解の場合は[NG]ボタンを押してください。
'
' 補足
' ・登録されているQuestionがランダムな順番で出題されます。
' ・1周目が終わると、2周目は間違ったものだけ出題されます、出題順は新たなランダム順です。
' ・全てが正解になるまで終わりません。がんばりましょう。
'
'-------------------------------------------------------------------------------
' グローバル変数の定義
'-------------------------------------------------------------------------------

' データが記載されている列名
Dim sColumnQs As String
Dim sColumnAs As String
Dim sColumnTy As String
Dim sColumnOK As String
Dim sColumnNG As String
Dim sColumnRT As String

' データ開始行
Dim nRowStart As Integer

' 1周目だけ Try/OK/NG をカウントアップするか、毎回カウントアップするか
Dim bUpdateOnlyFirstLoop As Boolean

' Excel内の処理対象シート
Dim TargetSheet As Worksheet


'-------------------------------------------------------------------------------
' 初期設定
'-------------------------------------------------------------------------------
Sub Init()

    ' データが記載されている列を設定
    sColumnQs = "A"  'Question列
    sColumnAs = "B"  'Answer列
    sColumnTy = "C"  'Try列
    sColumnOK = "D"  'OK列
    sColumnNG = "E"  'OK列
    sColumnRT = "F"  'Rate列
    
    'データ開始行 (1行目がヘッダ行の場合は2とする)
    nRowStart = 2

    ' 1周目だけ Try/OK/NG をカウントアップするか、毎回カウントアップするか
    bUpdateOnlyFirstLoop = True

    ' Excel内の処理対象シートを設定
    'Set TargetSheet = Worksheets(1)        'シート番号で指定
    Set TargetSheet = Worksheets("Sheet1") 'シート名で指定

End Sub


'-------------------------------------------------------------------------------
' 一次元配列の Debug Print
'-------------------------------------------------------------------------------
Sub DebugPlot_Matrix1D(indices)

    Debug.Print ("DebugPlot_Matrix1D")
    nSize = UBound(indices)
    Debug.Print "Size: " & nSize

    For i = 0 To UBound(indices)
        Debug.Print indices(i)
    Next i

End Sub


'-------------------------------------------------------------------------------
' 六次元配列の Debug Print
'-------------------------------------------------------------------------------
Sub DebugPlot_Matrix6D(elements)

    Debug.Print ("DebugPlot_Matrix6D")
    nSize = UBound(elements, 2)
    Debug.Print "Size: " & nSize

    For i = 0 To nSize - 1
        Debug.Print elements(0, i) & ", " & elements(1, i) & ", " & elements(2, i) & ", " & elements(3, i) & ", " & elements(4, i) & ", " & elements(5, i)
    Next i

End Sub


'-------------------------------------------------------------------------------
' 一次元配列のランダム並べ替え
'-------------------------------------------------------------------------------
Function RandomizeMatrix1D(indices)

    For i = 0 To UBound(indices)

        If indices(i) <> "-" Then
            Randomize
            rn = Int(UBound(indices) * Rnd)
            tmp = indices(i)
            indices(i) = indices(rn)
            indices(rn) = tmp
        End If

    Next i

    RandomizeMatrix1D = indices

End Function


'-------------------------------------------------------------------------------
' 一次元配列内の有効要素数をカウントする (ハイフン以外の要素数をカウントする)
'-------------------------------------------------------------------------------
Function CountValidElement(indices)

    CountValidElement = 0

    For i = 0 To UBound(indices)
        If indices(i) <> "-" Then
            CountValidElement = CountValidElement + 1
        End If
    Next i

End Function


'-------------------------------------------------------------------------------
' メイン
'-------------------------------------------------------------------------------
Sub Main()

    '-- 初期化 ---------------------------------------------
    Debug.Print "---Start---"
    Init
        
    '-- Question列の終了行を取得する --------------------------
    nRowEnd = Cells(Rows.Count, sColumnQs).End(xlUp).Row
    nRows = nRowEnd - nRowStart + 1
    'Debug.Print "nRowStart = " & nRowStart
    'Debug.Print "nRowEnd = " & nRowEnd
    'Debug.Print "nRows = " & nRows

    '-- Excel記載内容を配列に格納 -----------------------------
    Dim elements
    ReDim elements(6, nRows)
    For i = nRowStart To nRowEnd
        elements(0, i - nRowStart) = TargetSheet.Range(sColumnQs & i).Value
        elements(1, i - nRowStart) = TargetSheet.Range(sColumnAs & i).Value
        elements(2, i - nRowStart) = TargetSheet.Range(sColumnTy & i).Value
        elements(3, i - nRowStart) = TargetSheet.Range(sColumnOK & i).Value
        elements(4, i - nRowStart) = TargetSheet.Range(sColumnNG & i).Value
        elements(5, i - nRowStart) = TargetSheet.Range(sColumnRT & i).Value
    Next i

    '-- インデックス配列を作成 --------------------------------
    ReDim indices(nRows - 1)
    For i = 0 To UBound(indices)
        indices(i) = i
    Next i

    '-- Debug Print (開始時) -------------------------------
    DebugPlot_Matrix1D (indices)
    DebugPlot_Matrix6D (elements)

    '-- メインループ ----------------------------------------
    nloop = 0
    Do
        ' メインループ のループカウンターをカウントアップ
        nloop = nloop + 1

        ' カウンタ類を更新
        indices = RandomizeMatrix1D(indices) ' インデックス配列をランダムソート
        nValidQ = CountValidElement(indices) ' 有効問題数(未正解問題数)
        nAnswered = 0                        ' 回答済問題数

        '-- サブループ --
        For i = 0 To UBound(indices)

            ' インデックス(今回処理する行番号)を取得
            idx = indices(i)
            Debug.Print "MainLoop: " & nloop & "  SubLoop: " & i & "  Index: " & idx
            
            If idx <> "-" Then

                ' 回答済問題数をカウントアップ
                nAnswered = nAnswered + 1

                ' Excelデータ1行分を取得
                sQs = elements(0, idx)
                sAs = elements(1, idx)
                nTY = elements(2, idx)
                nOK = elements(3, idx)
                nNG = elements(4, idx)
                dRT = elements(5, idx)

                ' Try,OK,NG列が空欄の場合はゼロとする
                If nTY = "" Then nTY = 0
                If nOK = "" Then nOK = 0
                If nNG = "" Then nNG = 0

                ' Questionをメッセージボックスで表示
                sHeaderQ = "Lap " & nloop & " Question " & nAnswered & "/" & nValidQ
                rtnQ = MsgBox(sQs, vbOKOnly, sHeaderQ)
                If rtnQ = vbOK Then
                Else
                    Exit For
                End If

                ' Answerをメッセージボックスで表示
                sHeaderA = "Lap " & nloop & " Answer " & nAnswered & "/" & nValidQ
                rtnA = MsgBox(sAs, vbYesNo, sHeaderA)
                If rtnA = vbYes Then
                    nOK = nOK + 1
                    indices(i) = "-"  ' 正解したデータは、インデックスをハイフンに書き換える
                Else
                    nNG = nNG + 1
                End If
                nTY = nTY + 1
                dRT = nOK / nTY

                ' Try/OK/NG のカウントアップとExcelへの反映
                If ((bUpdateOnlyFirstLoop = True) And (nloop = 1)) Then

                    ' Update Matrix
                    elements(2, idx) = nTY
                    elements(3, idx) = nOK
                    elements(4, idx) = nNG
                    elements(5, idx) = dRT

                    ' Update Sheet
                    TargetSheet.Range(sColumnTy & (idx + nRowStart)).Value = elements(2, idx)
                    TargetSheet.Range(sColumnOK & (idx + nRowStart)).Value = elements(3, idx)
                    TargetSheet.Range(sColumnNG & (idx + nRowStart)).Value = elements(4, idx)
                    TargetSheet.Range(sColumnRT & (idx + nRowStart)).Value = elements(5, idx)
                End If

            End If

        Next i

    Loop While (CountValidElement(indices) <> 0) ' 全問正解になるまでループ(全要素がハイフンになるまでループ)
    
    '-- Debug Print (終了時) ----
    DebugPlot_Matrix1D (indices)
    DebugPlot_Matrix6D (elements)

    '-- End ---
    MsgBox "お疲れさまでした"
    Debug.Print "---Finish---"

End Sub

'-------------------------------------------------------------------------------
' EOC
