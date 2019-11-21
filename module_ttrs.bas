Option Explicit
#If Win64 Then
    Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Const ScorePrintCell      As String = "Q5"
'
Const ScorePlusCell      As String = "Q6"

Const ブロック開始セル  As String = "J4"
Const ゲーム開始セル    As String = "F5"
Const ゲーム開始行      As Integer = 5
Const ゲーム終了行      As Integer = 24
Const ゲーム開始列      As Integer = 6
Const ゲーム終了列      As Integer = 15
Const ブロック切取範囲  As String = "F6:O"
'Const ブロックの色      As Integer = 15
Const 壁底の色          As Integer = 1
Const ゲーム範囲列数    As Integer = 10
Const ゲーム開始行範囲  As String = "F5:O5"
Const ブロック塊の数    As Integer = 7
Const 上下に0マス       As Integer = 0
Const 左右に0マス       As Integer = 0
Const 下に1マス         As Integer = 1
Const 右に1マス         As Integer = 1
Const 左に1マス         As Integer = -1

Const speed0 As Variant = "00:00:01.00"
Const speed1 As Variant = "00:00:00.20"
Const speed2 As Variant = "00:00:00.30"
Const speed3 As Variant = "00:00:00.40"

Dim Check As Boolean

Dim 形                  As Integer
Dim 色                  As String
Dim ブロックの色      As Integer
Dim ブロックの向き      As String
Dim ブロックの向き2      As String

Dim fallCheck As Boolean
Dim speed As Variant

Dim timeNow As Variant
Sub Main()

    Call StartUpSetting

    Call MainGame

End Sub
Private Sub StartUpSetting()

    timeNow = Now() + Second("1")

    'Range(ScorePrintCell).NumberFormatLocal = "0_"
    Range(ScorePlusCell) = ""

    fallCheck = True
    speed = speed3

    Check = True
    ブロックの色 = 15

    Cells.Interior.ColorIndex = xlNone

    Cells.ColumnWidth = 2 '1.88

    With Range(ScorePrintCell)
    

        .Value = 0
        .Offset(0, 3).Value = "【↑】+【→】ブロックを右に90°回転させる。"
        .Offset(1, 3).Value = "【↑】+【←】ブロックを左に90°回転させる。"
        .Offset(2, 3).Value = "【↓】ブロックを下まで落とす。"
        .Offset(3, 3).Value = "【→】ブロックを右に動かす。"
        .Offset(4, 3).Value = "【←】ブロックを左に動かす。"
        .Offset(5, 3).Value = "【Enter】ゲーム終了。"
        .Offset(6, 3).Value = "【十字キー】と【Enter】キー以外使用禁止。"

    End With

    Call DrowGage
    Call tetorisu
    Application.Wait [Now() + "00:00:00.50"]
    Call zenkeshi
    
    Randomize

End Sub
Private Sub DrowGage()
    
    Dim i   As Long
    
    For i = ゲーム開始行 To ゲーム終了行 + 1
        Cells(i, ゲーム開始列 - 1).Interior.ColorIndex = 壁底の色
        Cells(i, ゲーム終了列 + 1).Interior.ColorIndex = 壁底の色
    Next i
    
    For i = ゲーム開始列 To ゲーム終了列
        Cells(ゲーム終了行 + 1, i).Interior.ColorIndex = 壁底の色
    Next i
    
End Sub
Private Sub MainGame()

    Do

        Call GameStart

        Do
            If GetAsyncKeyState(vbKeyReturn) Then
                End
            End If

            If GetAsyncKeyState(vbKeyDown) Then
                
                Check = False
                Do While ブロック同士のぶつかり判定("下") = False
                    Call ブロックを動かす(下に1マス, 左右に0マス)
                Loop
                Range(ScorePrintCell) = Range(ScorePrintCell).Value + 5
                Check = True
            End If


            If GetAsyncKeyState(vbKeyUp) Then
                If GetAsyncKeyState(vbKeyRight) Then
                    Call 右回転
                ElseIf GetAsyncKeyState(vbKeyLeft) Then
                    Call 左回転
                ElseIf (Range(ScorePrintCell).Value > 0) Then
                    Range(ScorePrintCell) = Range(ScorePrintCell).Value - 1
                    fallCheck = False
                End If
            End If

                
            If GetAsyncKeyState(vbKeyLeft) Then

                If ブロック同士のぶつかり判定("左") = False Then

                    Call ブロックを動かす(上下に0マス, 左に1マス)

                End If

            ElseIf GetAsyncKeyState(vbKeyRight) Then

                If ブロック同士のぶつかり判定("右") = False Then
            
                    Call ブロックを動かす(上下に0マス, 右に1マス)

                End If

            End If
            
            
            
            If ブロック同士のぶつかり判定("下") = False Then

                Call ブロックを動かす(下に1マス, 左右に0マス)

            Else
                
                If EndJudge = True Then
            
                    MsgBox "Score :  " & Range(ScorePrintCell).Value, Title:="Game Over"
                    End

                Else
                    
                    Call DestroyBlocks

                End If

                Exit Do

            End If

        



        Loop

    Loop

End Sub
Private Sub GameStart()

    ブロックの向き = "正面"

    Range(ブロック開始セル).Select

    形 = Int((ブロック塊の数 * Rnd) + 1)
    'ブロックの色 = Int((15 * Rnd + 1))

    With ActiveCell

        Select Case 形

            Case 1
                ブロックの色 = 41
                ' 2 1 3 4
                '■□■■ブロック
                Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(0, 2)).Select

            Case 2
                ブロックの色 = 43
                ' 1 2
                '□■
                '  ■■ブロック
                '   3 4
                Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 1), .Offset(1, 2)).Select

            Case 3
                ブロックの色 = 44
                ' 4
                '■
                '■□■ブロック
                ' 2 1 3
                Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(-1, -1)).Select

            Case 4
                ブロックの色 = 46
                ' 2 1 3
                '■□■
                '■    ブロック
                ' 4
                Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(1, -1)).Select

            Case 5
                ブロックの色 = 42
                '   1
                '  □
                '■■■ブロック
                ' 2 3 4
                Union(.Offset(0, 0), .Offset(1, -1), .Offset(1, 0), .Offset(1, 1)).Select

            Case 6
                ブロックの色 = 39
                '   3 4
                '  ■■
                '■□  ブロック
                ' 2 1
                Union(.Offset(0, 0), .Offset(0, -1), .Offset(-1, 0), .Offset(-1, 1)).Select

            Case 7
                ブロックの色 = 15
                ' 1 2
                '□■
                '■■ブロック
                ' 3 4
                Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 0), .Offset(1, 1)).Select

        End Select

    End With

End Sub
Private Function EndJudge() As Boolean

    Dim buf As Range

    EndJudge = False
    
    For Each buf In Range(ゲーム開始行範囲)

        If buf.Interior.ColorIndex <> xlNone Then

            EndJudge = True

        End If

    Next buf

End Function
Private Sub DestroyBlocks()
    Dim i                   As Integer
    Dim t                   As Integer
    Dim 色つきセル範囲の数  As Integer
    Dim bufAddress          As String
    Dim buf                 As Range
    
    
  
    Set buf = Range(ゲーム開始セル)

    For i = ゲーム開始行 To ゲーム終了行

        For t = ゲーム開始列 To ゲーム終了列
            
            Set buf = Union(buf, Cells(i, t))

            If Cells(i, t).Interior.ColorIndex <> xlNone Then

                色つきセル範囲の数 = 色つきセル範囲の数 + 1

            End If

        Next t
        
        If 色つきセル範囲の数 = ゲーム範囲列数 Then
        
                bufAddress = buf.Address
                Set buf = buf.Resize(buf.Rows.Count - 1)

                Cells(ゲーム終了行, ゲーム終了列).Select

                buf.Cut

                
                ActiveSheet.Paste Range(ブロック切取範囲 & buf.Rows.Count + ゲーム開始行)

                With Range(ScorePrintCell)
                    .Value = .Value + 100
                    .EntireColumn.AutoFit
                    DoEvents
                End With

                Set buf = Range(bufAddress)

        End If
        
    

        色つきセル範囲の数 = 0

    Next i

End Sub
Private Function ブロック同士のぶつかり判定(判定したい方向 As String) As Boolean

    ブロック同士のぶつかり判定 = False

    Select Case 判定したい方向

        Case "下"

            If 判定(下に1マス, 左右に0マス) = True Then

                ブロック同士のぶつかり判定 = True

            End If

        Case "右"

            If 判定(上下に0マス, 右に1マス) = True Then

                ブロック同士のぶつかり判定 = True

            End If

        Case "左"

            If 判定(上下に0マス, 左に1マス) = True Then

                ブロック同士のぶつかり判定 = True

            End If

    End Select

End Function
Private Function 判定(行 As Integer, 列 As Integer) As Boolean

    Dim buf1    As Range
    Dim buf2    As Range
    Dim fbuf    As Boolean

    判定 = False
    fbuf = False

    For Each buf1 In Selection
    
        For Each buf2 In Selection

            If buf1.Offset(行, 列).Address = buf2.Address Then

                fbuf = True

            End If

        Next buf2

        If fbuf = False Then

            If buf1.Offset(行, 列).Interior.ColorIndex <> xlNone Then

                判定 = True

                Exit Function

            End If

        End If

        fbuf = False

    Next buf1

End Function
Private Sub ブロックを動かす(行 As Integer, 列 As Integer)

    On Error Resume Next

    Selection.Interior.ColorIndex = xlNone

    'Call 回転
    
    Selection.Offset(行, 列).Interior.ColorIndex = ブロックの色

    Selection.Offset(行, 列).Select


'20->15->10
    If Check = True Then
        'Application.Wait [Now() + "00:00:00.20"]
                If fallCheck = False Then
                    'speed = speed0
                    Application.Wait [Now() + "00:00:01.00"]
                    fallCheck = True
                ElseIf Range(ScorePrintCell).Value >= 2000 Then
                    'speed = speed1
                    Application.Wait [Now() + "00:00:00.20"]
                ElseIf Range(ScorePrintCell).Value >= 1000 Then
                    'speed = speed2
                    Application.Wait [Now() + "00:00:00.30"]
                ElseIf True = True Then
                    'speed = speed3
                    Application.Wait [Now() + "00:00:00.40"]
                End If
    End If
    
    'Application.Wait [Now() + speed]
    'Application.Wait [Now() + "00:00:00.10"]
    
    On Error GoTo 0

End Sub
Private Sub 右回転()
    
    Dim 元の選択範囲    As Range
    Dim buf             As Range
    
        Set 元の選択範囲 = Selection
        Selection.Interior.ColorIndex = xlNone
        ブロックの向き2 = ブロックの向き
        ブロックの向き = ブロックを右に回転させる(形, ブロックの向き)

        For Each buf In Selection
            
            If buf.Interior.ColorIndex <> xlNone Then
                
                元の選択範囲.Select
                ブロックの向き = ブロックの向き2
                
                Exit For
                
            End If
            
        Next
        Selection.Interior.ColorIndex = ブロックの色

End Sub
Private Function ブロックを右に回転させる(形 As Integer, ブロックの向き As String) As String

    With ActiveCell

        Select Case 形

            Case 1
                '□■■■ブロック
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(-1, 1), .Offset(0, 1), .Offset(1, 1), .Offset(2, 1)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(1, -1), .Offset(1, 0), .Offset(1, 1), .Offset(1, 2)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(-1, 1), .Offset(0, 1), .Offset(1, 1), .Offset(2, 1)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(1, -1), .Offset(1, 0), .Offset(1, 1), .Offset(1, 2)).Select
                End Select

            Case 2
                ' 1 2
                '□■
                '  ■■ブロック
                '   3 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(0, 1), .Offset(1, 1), .Offset(0, 2), .Offset(-1, 2)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(0, -1), .Offset(0, 0), .Offset(1, 0), .Offset(1, 1)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(1, 1), .Offset(0, 2), .Offset(-1, 2)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(0, -1), .Offset(0, 0), .Offset(1, 0), .Offset(1, 1)).Select
                End Select

            Case 3
                ' 4
                '■
                '■□■ブロック
                ' 2 1 3
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(-1, 0), .Offset(-1, 1)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(1, 1)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(-1, 1), .Offset(1, 0), .Offset(1, 1)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(1, 0), .Offset(1, 1), .Offset(1, -1), .Offset(0, -1)).Select
                End Select

            Case 4
                ' 2 1 3
                '■□■
                '■    ブロック
                ' 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(-1, 0), .Offset(-1, -1)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(-1, 1)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(-1, 1), .Offset(1, 1), .Offset(1, 2)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(1, 0), .Offset(1, 1), .Offset(1, -1), .Offset(2, -1)).Select
                End Select

            Case 5
                '   1
                '  □
                '■■■ブロック
                ' 2 3 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(1, 1), .Offset(2, 0)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(1, -1), .Offset(0, 0), .Offset(0, -1), .Offset(0, -2)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(-1, 0), .Offset(-1, -1), .Offset(0, 0), .Offset(-2, 0)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(0, 0), .Offset(-1, 1), .Offset(0, 1), .Offset(0, 2)).Select
                End Select

            Case 6
                '   3 4
                '  ■■
                '■□  ブロック
                ' 2 1
                Select Case ブロックの向き
                Case "正面"
                    ブロックを右に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(-1, 0), .Offset(0, 1), .Offset(1, 1)).Select
                Case "左面"
                    ブロックを右に回転させる = "背面"
                    Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 0), .Offset(1, -1)).Select
                Case "背面"
                    ブロックを右に回転させる = "右面"
                    Union(.Offset(0, 0), .Offset(-1, 0), .Offset(0, 1), .Offset(1, 1)).Select
                Case "右面"
                    ブロックを右に回転させる = "正面"
                    Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 0), .Offset(1, -1)).Select
                End Select

            Case 7
                ' 1 2
                '□■
                '■■ブロック
                ' 3 4

        End Select

    End With

End Function
Private Sub 左回転()
    
    Dim 元の選択範囲    As Range
    Dim buf             As Range
    
        Set 元の選択範囲 = Selection
        Selection.Interior.ColorIndex = xlNone
        ブロックの向き2 = ブロックの向き
        ブロックの向き = ブロックを左に回転させる(形, ブロックの向き)

        For Each buf In Selection
            
            If buf.Interior.ColorIndex <> xlNone Then
                
                元の選択範囲.Select
                ブロックの向き = ブロックの向き2
                
                Exit For
                
            End If
            
        Next
        Selection.Interior.ColorIndex = ブロックの色

End Sub
Private Function ブロックを左に回転させる(形 As Integer, ブロックの向き As String) As String

    With ActiveCell

        Select Case 形

            Case 1
                '□■■■ブロック
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(-1, 1), .Offset(0, 1), .Offset(1, 1), .Offset(2, 1)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                    Union(.Offset(1, -1), .Offset(1, 0), .Offset(1, 1), .Offset(1, 2)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(-1, 1), .Offset(0, 1), .Offset(1, 1), .Offset(2, 1)).Select
                Case "右面"
                    ブロックを左に回転させる = "背面"
                    Union(.Offset(1, -1), .Offset(1, 0), .Offset(1, 1), .Offset(1, 2)).Select
                End Select

            Case 2
                ' 1 2
                '□■
                '  ■■ブロック
                '   3 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(1, 1), .Offset(0, 2), .Offset(-1, 2)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                    Union(.Offset(0, -1), .Offset(0, 0), .Offset(1, 0), .Offset(1, 1)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(0, 1), .Offset(1, 1), .Offset(0, 2), .Offset(-1, 2)).Select
                Case "右面"
                    ブロックを左に回転させる = "背面"
                    Union(.Offset(0, -1), .Offset(0, 0), .Offset(1, 0), .Offset(1, 1)).Select
                End Select

            Case 3
                ' 4
                '■
                '■□■ブロック
                ' 2 1 3
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(-1, 1), .Offset(1, 0), .Offset(1, 1)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                        Union(.Offset(1, 0), .Offset(1, 1), .Offset(1, -1), .Offset(0, -1)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(-1, 0), .Offset(-1, 1)).Select
                Case "右面"
                    ブロックを左に回転させる = "背面"
                                        Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(1, 1)).Select

                End Select

            Case 4
                ' 2 1 3
                '■□■
                '■    ブロック
                ' 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(0, 1), .Offset(-1, 1), .Offset(1, 1), .Offset(1, 2)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                    Union(.Offset(1, 0), .Offset(1, 1), .Offset(1, -1), .Offset(2, -1)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(-1, 0), .Offset(-1, -1)).Select
                    
                Case "右面"
                    ブロックを左に回転させる = "背面"
                    Union(.Offset(0, 0), .Offset(0, -1), .Offset(0, 1), .Offset(-1, 1)).Select
                    
                End Select

            Case 5
                '   1
                '  □
                '■■■ブロック
                ' 2 3 4
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(-1, 0), .Offset(-1, -1), .Offset(0, 0), .Offset(-2, 0)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                    Union(.Offset(0, 0), .Offset(-1, 1), .Offset(0, 1), .Offset(0, 2)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(1, 0), .Offset(1, 1), .Offset(2, 0)).Select
                    
                Case "右面"
                    ブロックを左に回転させる = "背面"
                    Union(.Offset(1, -1), .Offset(0, 0), .Offset(0, -1), .Offset(0, -2)).Select
                    
                End Select

            Case 6
                '   3 4
                '  ■■
                '■□  ブロック
                ' 2 1
                Select Case ブロックの向き
                Case "正面"
                    ブロックを左に回転させる = "右面"
                    Union(.Offset(0, 0), .Offset(-1, 0), .Offset(0, 1), .Offset(1, 1)).Select
                Case "左面"
                    ブロックを左に回転させる = "正面"
                    Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 0), .Offset(1, -1)).Select
                Case "背面"
                    ブロックを左に回転させる = "左面"
                    Union(.Offset(0, 0), .Offset(-1, 0), .Offset(0, 1), .Offset(1, 1)).Select
                Case "右面"
                    ブロックを左に回転させる = "背面"
                    Union(.Offset(0, 0), .Offset(0, 1), .Offset(1, 0), .Offset(1, -1)).Select
                End Select

            Case 7
                ' 1 2
                '□■
                '■■ブロック
                ' 3 4

        End Select

    End With

End Function


Sub tetorisu()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+o
'
    Range("I5:L5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("H7:M7").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K7:K8").Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I9:J9").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I11:I14").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("J12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I16:I17").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L16:L18").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("J19:K19").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I21:M21").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L22").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("K23").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("J24").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("L23").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("M24").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("AB16").Select
End Sub
Sub zenkeshi()
'
' Macro2 Macro
'

'
    Range("F5:O24").Select
    Range("O24").Activate
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("S13").Select
End Sub

