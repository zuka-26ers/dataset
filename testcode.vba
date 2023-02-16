
Public clickcounter As Long 'クリック回数をカウントする変数
Public UFL(1 To 5) As Boolean '左側に表示されるユーザーフォームの使用状況を示す配列
Public UFR(1 To 5) As Boolean '右側に表示されるユーザーフォームの使用状況を示す配列
Public ShapeList(1 To 5) As Shape 'ユーザーフォームに対応する図形のリスト
Dim obj As New Class1 'Class1のインスタンスを作成

Public Sub InitClass()
    Set obj.App = Application 'Class1のAppプロパティにApplicationオブジェクトを設定
'    Application.ActivePresentation.Save 'プレゼンテーションを保存（コメントアウトされている）
End Sub

Public Sub EndUF(uf As Variant) 'ユーザーフォームを終了するサブルーチン
    If Left(uf.Tag, 5) = "RIGHT" Then 'ユーザーフォームのタグがRIGHTで始まる場合
        UFR(CInt(Right(uf.Tag, 1))) = False 'タグの末尾の数字に対応するUFRの要素をFalseにする
    Else 'そうでない場合
        UFL(CInt(Right(uf.Tag, 1))) = False 'タグの末尾の数字に対応するUFLの要素をFalseにする
    End If
    ShapeList(CInt(uf.Caption)).Fill.Transparency = 1 'ユーザーフォームのキャプションに対応するShapeListの要素の塗りつぶしを透明にする
    Set ShapeList(CInt(uf.Caption)) = Nothing 'ShapeListの要素をNothingにする
    Unload uf 'ユーザーフォームをアンロードする
End Sub

Public Sub FormLimitter(ByVal UFName As String, oshp As Object) 'ユーザーフォームを表示するサブルーチン
    If oshp.Fill.Transparency = 1 Then '図形の塗りつぶしが透明なら
        Dim uf As Object 'ユーザーフォームの変数を宣言
        Set uf = VBA.UserForms.Add(UFName) 'ユーザーフォームを作成
        For i = 1 To 5 '1から5まで繰り返す
            If ShapeList(i) Is Nothing Then 'ShapeListの要素がNothingなら
                Set ShapeList(i) = oshp 'ShapeListの要素に図形を設定
                oshp.Fill.Transparency = 0.5 '図形の塗りつぶしの透過率を0.5にする
                uf.Caption = CStr(i) 'ユーザーフォームのキャプションにiを文字列に変換したものを設定
                If oshp.Left < 300 Then '図形の左端が300より小さいなら
                    For j = 1 To 5 '1から5まで繰り返す
                        If Not UFR(j) Then 'UFRの要素がFalseなら
                            UFR(j) = True 'UFRの要素をTrueにする
                            uf.Tag = "RIGHT" & CStr(j) 'ユーザーフォームのタグにRIGHTとjを文字列に変換したものを設定
                            uf.Left = 900 - j * 205 'ユーザーフォームの左端の位置を計算して設定
                            Exit For '繰り返しを抜け
                            End If
                    Next
                Else '図形の左端が300以上なら
                    For j = 1 To 5 '1から5まで繰り返す
                        If Not UFL(j) Then 'UFLの要素がFalseなら
                            UFL(j) = True 'UFLの要素をTrueにする
                            uf.Tag = "LEFT_" & CStr(j) 'ユーザーフォームのタグにLEFT_とjを文字列に変換したものを設定
                            uf.Left = -100 + j * 205 'ユーザーフォームの左端の位置を計算して設定
                            Exit For '繰り返しを抜ける
                        End If
                    Next
                End If
                uf.Show vbModeless 'ユーザーフォームをモードレスで表示
                Exit For '繰り返しを抜ける
            End If
        Next i
    End If
End Sub

Public Sub ボタン非表示()
    Dim sld As Slide
    Dim oshp As Shape
    For Each sld In ActivePresentation.Slides
        For Each oshp In sld.Shapes
            oshp.Line.Transparency = 1 '境界線色透過率(0->不透明・1->透明)
            If InStr(oshp.Name, "Action Button") > 0 Then
                oshp.TextFrame.TextRange.Text = "" '文字
                oshp.Fill.Transparency = 1 '図形塗りつぶしの透過率(0->不透明・1->透明)
            End If
        Next
    Next
End Sub

Public Sub ボタン表示()
    Dim sld As Slide
    Dim oshp As Shape
    For Each sld In ActivePresentation.Slides
        For Each oshp In sld.Shapes
            oshp.Line.Transparency = 0 '長方形色透過率(0->不透明・1->透明)
'            If InStr(oshp.Name, "Title") Then
'            End If
'            If InStr(oshp.Name, "Rectangle") Then
'            End If
'            If InStr(oshp.Name, "Rounded Rectangle") Then
'            End If
            If oshp.ActionSettings(ppMouseClick).Action = 8 Then
                oshp.Fill.ForeColor.RGB = RGB(255, 0, 0) '背景色
                With oshp.TextFrame.TextRange
                    .Font.Size = 10 '文字サイズ
                    .Font.Color.RGB = RGB(256, 200, 200) '文字色
                    .Text = Mid(oshp.ActionSettings(ppMouseClick).Run, 7) '関数名表示
'                    .Text = ""
                End With
            End If
        Next
    Next
End Sub

Public Sub 単独スライドテスト()
    Dim sld As Slide
    Dim oshp As Shape
    For Each sld In ActiveWindow.Selection.SlideRange
        For Each oshp In sld.Shapes
            oshp.Line.Transparency = 0 '長方形色透過率(0->不透明・1->透明)
'            If InStr(oshp.Name, "Title") Then
'            End If
'            If InStr(oshp.Name, "Rectangle") Then
'            End If
'            If InStr(oshp.Name, "Rounded Rectangle") Then
'            End If
            If oshp.ActionSettings(ppMouseClick).Action = 8 Then
                oshp.Fill.ForeColor.RGB = RGB(255, 0, 0) '背景色
                With oshp.TextFrame.TextRange
                    .Font.Size = 10 '文字サイズ
                    .Font.Color.RGB = RGB(256, 200, 200) '文字色
                    .Text = Mid(oshp.ActionSettings(ppMouseClick).Run, 7) '関数名表示
                    .Text = ""
                End With
            End If
        Next
    Next
End Sub

Public Sub popup_ポンプ(oshp As Object)
    FormLimitter "ポンプ", oshp
End Sub
Public Sub popup_バタフライ弁(oshp As Shape)
    FormLimitter "バタフライ弁", oshp
End Sub
Public Sub popup_空気作動弁(oshp As Shape)
    FormLimitter "空気作動弁", oshp
End Sub
Public Sub popup_ピストン弁(oshp As Shape)
    FormLimitter "ピストン弁", oshp
End Sub
Public Sub popup_空気弁ポジショナ(oshp As Shape)
    FormLimitter "空気弁ポジショナ", oshp
End Sub
Public Sub popup_バタフライ弁ポジショナ(oshp As Shape)
    FormLimitter "バタフライ弁ポジショナ", oshp
End Sub
Public Sub popup_電動弁(oshp As Shape)
    FormLimitter "電動弁", oshp
End Sub
Public Sub popup_電動バタフライ弁(oshp As Shape)
    FormLimitter "電動バタフライ弁", oshp
End Sub
Public Sub popup_ファン(oshp As Shape)
    FormLimitter "ファン_", oshp
End Sub
Public Sub popup_ダンパ(oshp As Shape)
    FormLimitter "ダンパ", oshp
End Sub
Public Sub popup_防火ダンパ(oshp As Shape)
    FormLimitter "防火ダンパ", oshp
End Sub
Public Sub popup_断路器(oshp As Shape)
    FormLimitter "断路器", oshp
End Sub
Public Sub popup_遮断器(oshp As Shape)
    FormLimitter "遮断器", oshp
End Sub
Public Sub popup_三方弁(oshp As Shape)
    FormLimitter "三方弁", oshp
End Sub
Public Sub popup_逆止弁(oshp As Shape)
    FormLimitter "逆止弁", oshp
End Sub
Public Sub popup_DG(oshp As Shape)
    FormLimitter "DG", oshp
End Sub

