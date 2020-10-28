# VBA Macro

- [VBA Macro](#vba-macro)
  - [Working the system as registration and list.](#working-the-system-as-registration-and-list)
    - [Source code](#source-code)
    - [Clarification](#clarification)

## Working the system as registration and list.

### Source code

```text

Sub register_click()
    If Worksheets("Register").Range("C2").Value = "" And Worksheets("Register").Range("C4").Value = "" Then
        MsgBox "Please type in the textbox.", vbOKOnly, "Message"
        Exit Sub
    Else
        Dim inputRow As Long
        With Worksheets("List")
                inputRow = .Range("A1").CurrentRegion.Rows.Count + 1
                .Cells(inputRow, 1).Value = _
                    WorksheetFunction.Max(.Range("A:A")) + 1
                .Cells(inputRow, 2).Value = Date
                .Cells(inputRow, 2).NumberFormat = "yyyy-mm-dd"
                .Cells(inputRow, 3).Value = Range("C2").Value
                .Cells(inputRow, 4).Value = Range("C4").Value
                .Cells(inputRow, 1).Borders.LineStyle = xlContinuous
                .Cells(inputRow, 2).Borders.LineStyle = xlContinuous
                .Cells(inputRow, 3).Borders.LineStyle = xlContinuous
                .Cells(inputRow, 4).Borders.LineStyle = xlContinuous
                .Cells(inputRow, 1).Borders.Weight = xlMedium
                .Cells(inputRow, 2).Borders.Weight = xlMedium
                .Cells(inputRow, 3).Borders.Weight = xlMedium
                .Cells(inputRow, 4).Borders.Weight = xlMedium
        End With
        Range("C2:C4") = ""
        MsgBox "Completed the registration.", vbOKOnly, "Completed"
    End If
End Sub

```

### Clarification

  IfステートメントでRegisterシートの入力欄であるC2セルとC4セルの中にテキストの情報が入っているかいないかで分岐をしてから処理を書いていく。

- もし、RegisterシートのC2とC4に値が入っていない（””で空っぽの意味）場合、メッセージボックスにて、テキストボックスの中に入力をしてくださいというメッセージを表示させて、そのまま処理は終了する（Exit Sub）。
- もし、RegisterシートのC2とC4に値が入っている（Else以降）場合、まず変数inputRowを定義する。
    - inputRowはListシートで新たに情報が追加される行のことを示している。
    - inputRowの1列目に入る数値は、実際の上から数えられている数字より1つ大きい
      - なぜ大きいかというと、1行目には番号、名前、チームと言った見出しが書かれているため、始点位置が1つ下にずれているから忘れずに+1を書く。
    - inputRowの2列目に入る情報としては、その情報が登録された（入力された）日付が入る。
      - "yyyy/mm/dd"で処理を行おうとすると、文字列として認識されてしまっているのか「#######」と表示されてしまうため、”yyyy-mm-dd”で表示設定を行う
    - inputRowの3列目と4列目に入る情報は、RegisterシートのC2とC4に入力された情報を転記してListシートに反映する。
    - 入力されたセルが追加されていくごとに、罫線（黒い線の□）を追加していくために、"Borders.LineStyle = xlContinuous"を使う。
    - また、罫線の太さを見出しの罫線と合わせるために、"Borders.Weight = xlMedium"を使う。
    - "Range（”C2:C4") = """でRegisterシート内に入力されていた情報を空っぽにする。
    - 登録が完了した旨を、メッセージボックスで表示して知らせる。