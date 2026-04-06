Attribute VB_Name = "Module1"
' ============================================================
'  履歴書自動生成マクロ
'  Excel入力シート → Wordテンプレート → 完成Word出力
' ============================================================
' 【使い方】
'   1. このコードをExcelのVBAエディタ（Alt+F11）に貼り付ける
'   2. 参照設定で「Microsoft Word xx.x Object Library」を有効にする
'      （ツール → 参照設定 → Microsoft Word にチェック）
'   3. 設定シートのパスを自分の環境に合わせて変更する
'   4. マクロ「履歴書生成」を実行する
' ============================================================

Option Explicit

Sub 履歴書生成()

    ' ---------- 設定シートからパス取得 ----------
    Dim ws設定 As Worksheet
    Set ws設定 = ThisWorkbook.Sheets("設定")

    Dim templatePath As String
    Dim outputFolder As String
    Dim outputFileName As String

    templatePath = ws設定.Range("B2").Value
    outputFolder = ws設定.Range("B3").Value
    outputFileName = ws設定.Range("B4").Value

    ' パスの末尾に円マークを保証
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"

    ' ---------- パス存在チェック ----------
    If templatePath = "" Or Dir(templatePath) = "" Then
        MsgBox "Wordテンプレートが見つかりません。" & vbCrLf & templatePath, vbCritical
        Exit Sub
    End If
    If outputFolder = "\" Or Dir(outputFolder, vbDirectory) = "" Then
        MsgBox "出力先フォルダが見つかりません。" & vbCrLf & outputFolder, vbCritical
        Exit Sub
    End If

    ' ---------- データ読み込み ----------
    Dim ws基本 As Worksheet
    Dim ws歴   As Worksheet
    Dim ws資格 As Worksheet
    Set ws基本 = ThisWorkbook.Sheets("基本情報")
    Set ws歴 = ThisWorkbook.Sheets("学歴職歴")
    Set ws資格 = ThisWorkbook.Sheets("資格")

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim i As Integer
    For i = 2 To ws基本.Cells(ws基本.Rows.Count, 1).End(xlUp).Row
        Dim key As String
        Dim val As String
        key = ws基本.Cells(i, 1).Value
        val = CStr(ws基本.Cells(i, 2).Value)
        Select Case key
            Case "氏名（漢字）":    dic("{{氏名}}") = val
            Case "ふりがな":        dic("{{ふりがな}}") = val
            Case "生年月日":        dic("{{生年月日}}") = val
            Case "年齢":            dic("{{年齢}}") = val
            Case "性別":            dic("{{性別}}") = val
            Case "現住所〒":        dic("{{現住所郵便番号}}") = val
            Case "現住所":          dic("{{現住所}}") = val
            Case "現住所ふりがな":  dic("{{現住所ふりがな}}") = val
            Case "電話番号":        dic("{{電話}}") = val
            Case "携帯番号":        dic("{{携帯}}") = val
            Case "FAX":             dic("{{FAX}}") = val
            Case "メールアドレス":  dic("{{メール}}") = val
            Case "作成年":          dic("{{作成年}}") = val
            Case "作成月":          dic("{{作成月}}") = val
            Case "作成日":          dic("{{作成日}}") = val
            Case "本人希望":        dic("{{本人希望}}") = val
            Case "障がいの状況":    dic("{{障がい状況}}") = val
            Case "自己PR":          dic("{{自己PR}}") = val
        End Select
    Next i

    ' 学歴職歴（最大20行）
    Dim maxReki As Integer
    maxReki = 20
    For i = 1 To maxReki
        Dim rowIdx  As Integer
        Dim nen     As String
        Dim tuki    As String
        Dim naiyou  As String
        rowIdx = i + 1
        If rowIdx <= ws歴.Cells(ws歴.Rows.Count, 4).End(xlUp).Row Then
            nen = CStr(ws歴.Cells(rowIdx, 1).Value)
            tuki = CStr(ws歴.Cells(rowIdx, 2).Value)
            naiyou = Trim(ws歴.Cells(rowIdx, 3).Value & " " & ws歴.Cells(rowIdx, 4).Value)
        Else
            nen = "": tuki = "": naiyou = ""
        End If
        dic("{{歴年" & i & "}}") = nen
        dic("{{歴月" & i & "}}") = tuki
        dic("{{歴内容" & i & "}}") = naiyou
    Next i

    ' 資格（最大4行）
    Dim maxShikaku As Integer
    maxShikaku = 4
    For i = 1 To maxShikaku
        Dim snen  As String
        Dim stuki As String
        Dim sname As String
        rowIdx = i + 1
        If rowIdx <= ws資格.Cells(ws資格.Rows.Count, 3).End(xlUp).Row Then
            snen = CStr(ws資格.Cells(rowIdx, 1).Value)
            stuki = CStr(ws資格.Cells(rowIdx, 2).Value)
            sname = ws資格.Cells(rowIdx, 3).Value
        Else
            snen = "": stuki = "": sname = ""
        End If
        dic("{{資格年" & i & "}}") = snen
        dic("{{資格月" & i & "}}") = stuki
        dic("{{資格内容" & i & "}}") = sname
    Next i

    ' ---------- Wordを起動してテンプレートをコピー ----------
    Dim wdApp As Object
    Dim wdDoc As Object

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    wdApp.Visible = False

    Dim outputPath As String
    outputPath = outputFolder & outputFileName & ".docx"

    If Dir(outputPath) <> "" Then
        Dim ans As Integer
        ans = MsgBox("同名ファイルが存在します。上書きしますか？" & vbCrLf & outputPath, _
                     vbYesNo + vbQuestion)
        If ans = vbNo Then
            wdApp.Quit
            Exit Sub
        End If
        Kill outputPath
    End If

    FileCopy templatePath, outputPath
    Set wdDoc = wdApp.Documents.Open(outputPath)

    ' ---------- プレースホルダー一括置換 ----------
    Dim ph As Variant
    For Each ph In dic.Keys
        Call ReplaceInDoc(wdDoc, CStr(ph), CStr(dic(ph)))
    Next ph

    ' ---------- 顔写真を挿入 ----------
    Dim photoPath As String
    photoPath = ws設定.Range("B5").Value
    If photoPath <> "" And Dir(photoPath) <> "" Then
        Call 顔写真を挿入(wdDoc, photoPath)
    End If
    
    ' ---------- 保存して閉じる ----------
    wdDoc.Save
    wdDoc.Close
    wdApp.Quit

    MsgBox "履歴書を出力しました！" & vbCrLf & outputPath, vbInformation, "完了"
    Shell "explorer.exe """ & outputFolder & """", vbNormalFocus

End Sub

' ============================================================
'  プレースホルダー置換サブルーチン
' ============================================================
Private Sub ReplaceInDoc(ByVal doc As Object, _
                         ByVal findStr As String, _
                         ByVal replaceStr As String)
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findStr
        .Replacement.Text = replaceStr
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = True
        .Execute Replace:=2
    End With
End Sub

' ============================================================
'  顔写真挿入サブルーチン（座標指定版）
'  ブックマーク「写真欄」に挿入し、点線枠にぴったり合わせる
' ============================================================
Private Sub 顔写真を挿入(ByVal doc As Object, ByVal photoPath As String)

    On Error GoTo ErrHandler

    Dim shp As Object

    ' ブックマーク「写真欄」が設定済みの場合はそちらを優先
    If doc.Bookmarks.Exists("写真欄") Then
        Dim rng As Object
        Set rng = doc.Bookmarks("写真欄").Range
        Set shp = doc.InlineShapes.AddPicture( _
            Filename:=photoPath, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=rng)

        ' 標準規格（縦40mm×横30mm）に変更する場合
        shp.LockAspectRatio = False
        shp.Width = 2.835 * 30    ' 横30mm = 85.05pt
        shp.Height = 2.835 * 40   ' 縦40mm = 113.4pt
    
    Else
        ' ブックマークがない場合は絶対座標で配置
        ' Left=4228465EMU=333.3pt, Top=114300EMU=9pt
        Set shp = doc.Shapes.AddPicture( _
            Filename:=photoPath, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Left:=333.3, _
            Top:=9, _
            Width:=85.05, _
            Height:=113.4)
        shp.ZOrder 0
    End If

    Exit Sub

ErrHandler:
    MsgBox "写真の挿入に失敗しました。" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================
'  テンプレートを選択
' ============================================================
Sub テンプレートを選択()
    Dim path As String
    
    ' ファイル選択ダイアログを開く
    path = Application.GetOpenFilename( _
        FileFilter:="Wordファイル (*.docx;*.doc),*.docx;*.doc", _
        Title:="Wordテンプレートを選択してください")
    
    ' キャンセルされた場合は何もしない
    If path = "False" Then Exit Sub
    
    ' B2セルにパスを書き込む
    ThisWorkbook.Sheets("設定").Range("B2").Value = path
    
    MsgBox "テンプレートを設定しました！" & vbCrLf & path, vbInformation
End Sub

' ============================================================
'  出力フォルダを選択
' ============================================================
Sub 出力フォルダを選択()
    Dim path As String
    
    ' フォルダ選択ダイアログを開く
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "出力先フォルダを選択してください"
        .AllowMultiSelect = False
        
        ' キャンセルされた場合は何もしない
        If .Show = False Then Exit Sub
        
        path = .SelectedItems(1)
    End With
    
    ' B3セルにパスを書き込む
    ThisWorkbook.Sheets("設定").Range("B3").Value = path
    
    MsgBox "出力先フォルダを設定しました！" & vbCrLf & path, vbInformation
End Sub

' ============================================================
'  写真を選択
' ============================================================
Sub 写真を選択()
    Dim path As String
    path = Application.GetOpenFilename( _
        FileFilter:="画像ファイル (*.jpg;*.jpeg;*.png;*.bmp),*.jpg;*.jpeg;*.png;*.bmp", _
        Title:="顔写真を選択してください")
    If path = "False" Then Exit Sub
    ThisWorkbook.Sheets("設定").Range("B5").Value = path
    MsgBox "写真を設定しました！", vbInformation
End Sub
