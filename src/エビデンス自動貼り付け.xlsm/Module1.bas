Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Sub Kicker()
    Call AutoCapture
End Sub

Sub AutoCapture()

Dim CB As Variant
Dim j As Integer: j = 0
Dim ShapesCount As Shape
Dim ShapesCounts As Long: ShapesCounts = 0

'おまじない
On Error Resume Next
Application.ScreenUpdating = False

'タイトルの文字列に"停止中"の文字列がある場合、Quitへ
If Right(Application.Caption, 3) = "停止中" Then
    GoTo Quit
End If

'タイトルに"実行中"の文字列を挿入
Application.Caption = "Capture実行中"

CB = Application.ClipboardFormats
   
If CB(1) <> -1 Then
    For i = 1 To UBound(CB)
        'クリップボードがビットマップ形式の場合以下を実行
        If CB(i) = xlClipboardFormatBitmap Then
            
            '"エビデンス"シートに対して以下を実行
            With ThisWorkbook.Sheets("エビデンス")
                .Select
                j = 1
                
                'シートないのオブジェクトの数が0の場合、「ShapesCounts」を2
                If .Shapes.Count = 0 Then
                    ShapesCounts = 2
                Else
                    'シート内のオブジェクトの一番最後のセル番号を求める
                    For Each ShapesCount In .Shapes
                        ShapesCounts = Application.Max(ShapesCounts, ShapesCount.BottomRightCell.Row)
                        j = j + 1
                    Next
                    ShapesCounts = ShapesCounts + 3
                End If
                
                '貼り付け位置に移動
                Application.Goto reference:=ThisWorkbook.Sheets("エビデンス").Range("B" & ShapesCounts), scroll:=True
                
                '画面を最小化（上記スクロール処理を実行するとこのブックが最前に来てしまうため）
                Application.WindowState = xlMinimized
                
                '移動した位置に画像を貼り付け
                .Paste Destination:=ActiveCell.Offset(1, 1)
                Selection.ShapesRange.Line.Visible = msoTrue
                
                'インクリメントした番号を記載
                'Range("B" & ShapesCounts).Value = "#" & j
                Range("B" & ShapesCounts).Formula = "=""#"" & COUNTA($B$1:B" & ShapesCounts - 1 & ") + 1"
            End With
                            
            'クリップボードを空にする
            OpenClipboard
            EmptyClipboard
            CloseClipboard
        End If
    Next i
End If
DoEvents

'1秒ごとに自身を呼ぶ
Application.OnTime DateAdd("s", 1, Now), "AutoCapture"
Exit Sub

Quit:
'タイトルの文字列を消去する
Application.Caption = ""
End Sub

Sub cahngesheets()

Dim AdShName As String
Dim Exculusion() As Variant
ReDim Exculusion(0)
Exculusion = Array(":", "\", "/", "?", "*", "[", "]")

Dim FromatDate As String

'おまじない
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
On Error Resume Next

'タイトルの文字列を"停止中"にする
Application.Caption = "停止中"

If Not ThisWorkbook.Sheets("エビデンス").Shapes.Count = 0 Then
    'シート名用の現在時刻を取得
    FromatDate = Format(Now, "yy年mm月dd日 hh時mm分ss秒")
    
    'シートを成形
    With ThisWorkbook.Sheets("エビデンス")
        .Active
        ActiveWindow.Zoom = 60
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        Range("B2").Select
        '"エビデンス"シートをシートの最後に移動
        .Move after:=Worksheets(Worksheets.Count)
        
        With Worksheets(Worksheets.Count)
        '移動したシートの名前を変更
            AdShName = InputBox("シート名を入力してください", "シート名")
            If AdShName = "" Then
                .Name = FromatDate
            Else
                If Len(AdShName) >= 32 Then
                    MsgBox ("シート名が長すぎます")
                    .Name = FromatDate
                Else
                    For Each ExculusionList In Exculusion
                        If InStr(AdShName, ExculusionList) Then
                            MsgBox ("使用できない文字列が含まれています")
                            .Name = FromatDate
                            Exit For
                        End If
                    Next
                End If
            Call CreateNewWorksheet(AdShName)
            End If
        End With
    End With
    
    'シートの前から2番目に"エビデンス"シートを追加
    Worksheets.Add after:=Worksheets("ツール実行")
    ActiveSheet.Name = "エビデンス"
    
    'エビデンスシートをフォーマットに整形
    Cells.Font.Name = "ＭＳ Ｐゴシック"
    Columns("A").ColumnWidth = 4
    Columns("B").ColumnWidth = 4
    ActiveWindow.Zoom = 60
    Columns("B").Font.Name = "ＭＳ ゴシック"
    Columns("B").Font.Size = 10
End If

'先頭のシートをアクティブにする
Sheets(1).Activate
Range("A1").Select

'おまじない解除
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub sheetsdelete()

Dim Bookname As String
Dim CopyBook As String

Dim CopySh As Worksheet
Dim ArrayShName() As String
ReDim ArrayShName(0)

Dim i As Long: i = 0
Dim j As Long: j = 0

Dim delSh As Long
Dim delSh_CNT As Long

'おまじない
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
On Error Resume Next
    
'保存対象のシートを検索
For Each CopySh In ThisWorkbook.Worksheets
    If Not (CopySh.Name = "ツール実行" Or CopySh.Name = "エビデンス") Then
        ReDim Preserve ArrayShName(i)
        ArrayShName(i) = CopySh.Name
        i = i + 1
    End If
Next CopySh

If Not i = 0 Then
    'ブック名指定
    Bookname = Format(Now(), "エビデンス_yy年mm月dd日 hh時mm分ss秒") & ".xlsx"
    CopyBook = ActiveWorkbook.Path & "\" & Bookname
    
    'シートを選択して新しいブックへコピー
    Worksheets(ArrayShName).Select
    Worksheets(ArrayShName).Copy
    
    '表示のサイズなどを統一
    For j = 1 To Sheets.Count
        Sheets(Sheets(j).Name).Select
        Cells.Font.Name = "ＭＳ Ｐゴシック"
        Columns("A").ColumnWidth = 4
        Columns("B").ColumnWidth = 4
        ActiveWindow.Zoom = 60
        Columns("B").Font.Name = "ＭＳ ゴシック"
        Columns("B").Font.Size = 10
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        Range("B2").Select
    Next j
    
    '先頭のシートをアクティブにする
    Sheets(1).Activate
    
    '上記で設定した名前で保存して閉じる
    ActiveWorkbook.SaveAs Filename:=CopyBook
    ActiveWorkbook.Close
    
    '「エビデンス」シートより右のシートを全て削除
    delSh_CNT = Worksheets("エビデンス").Index + 1
    
    For delSh = Worksheets.Count To delSh_CNT Step -1
        Worksheets(delSh).Delete
    Next
    
    '先頭のシートをアクティブにする
    Sheets(1).Activate
    Range("A1").Select
    
    'おまじないを解除して、本ブックにに上書き処理をかける
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End If
    
End Sub

Sub macrostop()
    Application.Caption = "停止中"
    '先頭のシートをアクティブにする
    Sheets(1).Activate
    Range("A1").Select
End Sub

'WorkSheet名を与えることでワークシートを新規作成
'重複したworksheetがある場合、(1), (2), ...と連番になる。
'呼び出し側には作成したワークシート名(ByRef)を返す。
Sub CreateNewWorksheet(ByRef SheetName As String)
    
    Dim i As Long, j As Long
    Dim rc As String, tmpSN As String
    Dim NewMode As Boolean
    Dim CworkS As Boolean 'Loopの終了判定(Checkworksheet)
    CworkS = True
    tmpSN = SheetName

        '新しいグラフ描画用シートの作成
        Do While CworkS
            CworkS = WorkSheetCheck(tmpSN)
            '重複したworksheetがある場合、(1), (2), ...と連番をつけて,
            '更に重複が無いか調べてない場合はWorksheet名として適用する。
            j = j + 1 '連番用の変数
            If CworkS Then
                'worksheet新規作成モードに切り替え
                NewMode = True
                'シート名の再設定
                tmpSN = SheetName & "(" & j & ")"
            End If
        Loop
        
        'ワークシートを最後尾に新規作成し、指定したファイル名にする。
        If NewMode Then
            Worksheets(Worksheets.Count).Name = tmpSN
        Else
            Worksheets(Worksheets.Count).Name = SheetName
        End If

End Sub


'重複したWorksheetが有るかチェックする。
'引数;検索するシート名 CheckSheets
'戻り値; 重複検出=>WSC=True, 重複なし=>WSC=False
Function WorkSheetCheck(ByVal CheckSheets) As Boolean
    Dim i As Long
    WorkSheetCheck = False
    Dim tmpShChar As String, tmpChar As String
    
    
    For i = 1 To Worksheets.Count
        'シート名は大文字小文字区別されないのでここで、全て小文字に変換しておく。
        tmpShChar = LCase(Worksheets(i).Name)
        tmpChar = LCase(CheckSheets)
        
        tmpChar = StrConv(tmpChar, vbWide)
        tmpChar = StrConv(tmpChar, vbHiragana)
        tmpShChar = StrConv(tmpShChar, vbWide)
        tmpShChar = StrConv(tmpShChar, vbHiragana)
        
        
        If tmpShChar = tmpChar Then
            WorkSheetCheck = True
            Exit Function
        End If
    Next
    
End Function

