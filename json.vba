''
' VBA json parse v1.0.0
' (c) BIWAZ DESIGN - Takeshi Matsui - https://github.com/biwaz-design/opencode/blob/main/json.vba
'
' It's a completely original JSON parser, but I think the structure will not
' change much no matter who makes it. I wanted to improve the speed of the
' off-the-shelf parser as much as possible, so I devised it. I hope it will be
' useful for your work.
' -----------------------------------------------------------------------
' 完全にオリジナルのJSONパーサーですが、誰が作っても構造はあまり変わらないと思います。
' 既製のパーサーの速度をできるだけ向上させたいと思ったので、そこのところ頑張りました。
' お役に立てば幸いです。
'
' * Parse json-string to object
' * Stringify object to json-string
'
' Errors:
' 32000 - json parse
'
' @class JsonConverter
' @author biwaz-design@outlook.jp
' @license MIT (http://www.openbiwaz.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Private biwaz, designs, off, idx, whitespace

Private Function StringifyTab(obj, ByVal off)
    Dim i, tp, sep, ary, key, data
    tp = VarType(obj)
    Select Case tp
    Case vbString
        StringifyTab = """" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(obj, "\", "\\"), Chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), "/", "\/") & """"
    Case vbObject
        sep = vbCrLf & String(off + 1, vbTab)
        Select Case TypeName(obj)
        Case "Dictionary"
            If 0 < obj.Count Then
                ReDim ary(obj.Count - 1)
                i = 0
                If IsNull(whitespace) Then
                    For Each key In obj.Keys
                        ary(i) = sep + StringifyTab(key, off + 1) & ":" & StringifyTab(obj(key), off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & "}"
                Else
                    sep = vbCrLf & String(off + 1, whitespace)
                    For Each key In obj.Keys
                        ary(i) = sep + StringifyTab(key, off + 1) & ": " & StringifyTab(obj(key), off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & Left(sep, Len(sep) - 1) & "}"
                End If
                ary(0) = "{" & ary(0)
                StringifyTab = Join(ary, ",")
            Else
                StringifyTab = "{}"
            End If
        Case "Collection"
            If 0 < obj.Count Then
                ReDim ary(obj.Count - 1)
                i = 0
                If IsNull(whitespace) Then
                    For Each data In obj
                        ary(i) = StringifyTab(data, off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & "]"
                Else
                    sep = vbCrLf & String(off + 1, whitespace)
                    For Each data In obj
                        ary(i) = sep & StringifyTab(data, off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & Left(sep, Len(sep) - 1) & "]"
                End If
                ary(0) = "[" & ary(0)
                StringifyTab = Join(ary, ",")
            Else
                StringifyTab = "[]"
            End If
        End Select
    Case vbNull
        StringifyTab = "null"
    Case vbBoolean
        If obj Then StringifyTab = "true" Else StringifyTab = "false"
    Case vbDate
        StringifyTab = """" & obj & """"
    Case Else
        If vbArray <= tp And tp <= vbArray + vbByte Then
            If 0 < UBound(obj) Then
                ReDim ary(UBound(obj))
                i = 0
                If IsNull(whitespace) Then
                    For Each data In obj
                        ary(i) = StringifyTab(data, off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & "]"
                Else
                    sep = vbCrLf & String(off + 1, whitespace)
                    For Each data In obj
                        ary(i) = sep & StringifyTab(data, off + 1)
                        i = i + 1
                    Next
                    ary(UBound(ary)) = ary(UBound(ary)) & Left(sep, Len(sep) - 1) & "]"
                End If
                ary(0) = "[" & ary(0)
                StringifyTab = Join(ary, ",")
            Else
                StringifyTab = "[]"
            End If
        Else
            StringifyTab = obj
        End If
    End Select
End Function

Public Function Stringify(obj, Optional ws)
    If IsMissing(ws) Then whitespace = Null Else whitespace = ws
    If IsNull(whitespace) Then
        Stringify = Replace(Replace(Replace(Replace(Replace(Replace(Replace(StringifyTab(obj, 0), "\", "\\"), Chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), "/", "\/")
    Else
        Stringify = StringifyTab(obj, 0)
    End If
    Dim i
    For Each i In Array(0, 1, 2, 3, 4, 5, 6, 7, 11, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
        Stringify = Replace(Stringify, Chr(i), "\u00" & Right("0" & Hex(i), 2))
    Next
End Function

Private Sub ParseCore(ByRef Value)
    Dim ch, child

    ch = Mid(biwaz, off, 1)
    Select Case ch
    Case "{"
        Set Value = CreateObject("Scripting.Dictionary")
        off = off + 1
        ch = Mid(biwaz, off, 1)
        If ch = "}" Then
            off = off + 1
            Exit Sub
        End If

        Do
            If ch <> """" Then Err.Raise 32000, "json parse", "オブジェクトのキーが検出できません" ' Unable to find key ob object
            off = off + 1
            If Mid(biwaz, off, 1) <> ":" Then Err.Raise 32000, "json parse", "オブジェクトのキー" & designs(idx) & "の次に ':' を検出できません" ' Unable to find ':' next to object key & designs(idx)
            ch = designs(idx)
            idx = idx + 1

            off = off + 1
            ParseCore child
            If VarType(child) = 9 Then Set Value(ch) = child Else Value(ch) = child

            child = ch
            ch = Mid(biwaz, off, 1)
            If ch = "}" Then
                off = off + 1
                Exit Sub
            ElseIf ch <> "," Then
                Err.Raise 32000, "json parse", "オブジェクトのメンバー """ & child & """:.. の次に ',' を検出できません" ' Unable to find ',' next to object member & child & : ..
            End If

            off = off + 1
            ch = Mid(biwaz, off, 1)
        Loop
    Case "["
        Set Value = New Collection
        off = off + 1
        ch = Mid(biwaz, off, 1)
        If ch = "]" Then
            off = off + 1
            Exit Sub
        End If

        ParseCore child
        Value.Add child
        
        Do
            ch = Mid(biwaz, off, 1)
            If ch = "]" Then
                off = off + 1
                Exit Sub
            ElseIf ch <> "," Then
                Err.Raise 32000, "json parse", "配列要素の次に ',' を検出できません" ' Unable to find ',' next to array element
            End If
            off = off + 1
            
            ParseCore child
            Value.Add child
        Loop
    Case """"
        off = off + 1
        Value = designs(idx)
        idx = idx + 1
    Case "t"
        If Mid(biwaz, off, 4) <> "true" Then Err.Raise 32000, "json parse", "'t' の次に 'rue' が検出できません" ' 'rue' cannot be detected after 't'
        off = off + 4
        Value = True
    Case "f"
        If Mid(biwaz, off, 5) <> "false" Then Err.Raise 32000, "json parse", "'f' の次に 'alse' が検出できません" ' 'alse' cannot be detected after 'f'
        off = off + 5
        Value = False
    Case "n"
        If Mid(biwaz, off, 4) <> "null" Then Err.Raise 32000, "json parse", "'n' の次に 'ull' が検出できません" ' 'ull' cannot be detected after 'n'
        off = off + 4
        Value = Null
    Case Else
        Dim length, number, ac
        length = Len(biwaz)
        If ch = "-" Then
            off = off + 1
            If length < off Then Err.Raise 32000, "json parse", "数値が記号 - の後、途切れています" ' The number is broken after the symbol-
            number = ch
            ch = Mid(biwaz, off, 1)
        Else
            number = ""
        End If

        ' integer
        off = off + 1
        number = number + ch
        ac = Asc(ch)
        If 48 < ac And ac < 58 Then
            Do Until length < off
                ch = Mid(biwaz, off, 1)
                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Exit Do
                off = off + 1
                number = number + ch
            Loop
        ElseIf ac <> 48 Then
            Err.Raise 32000, "json parse", "不明なトークンです (" & number & ")" ' Unknown token ( & number & )
        End If

        ' fraction
        If off <= length Then
            ch = Mid(biwaz, off, 1)
            If ch = "." Then
                off = off + 1
                number = number + ch
                If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & number & ")" ' The numbers are interrupted in the middle ( & number & )

                ch = Mid(biwaz, off, 1)
                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & number & ")" ' The numbers are interrupted in the middle ( & number & )

                Do
                    off = off + 1
                    number = number + ch
                    If length < off Then Exit Do
                    ch = Mid(biwaz, off, 1)
                    ac = Asc(ch)
                Loop Until ac < 48 Or 58 <= ac
            End If
        End If

        ' exponent
        If off <= length Then
            Select Case ch
            Case "E", "e"
                off = off + 1
                number = number + ch
                If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & number & ")" ' The numbers are interrupted in the middle ( & number & )

                ch = Mid(biwaz, off, 1)
                Select Case ch
                Case "-", "+"
                    off = off + 1
                    number = number + ch
                    If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & number & ")" ' The numbers are interrupted in the middle ( & number & )
                    ch = Mid(biwaz, off, 1)
                End Select

                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & number & ")" ' The numbers are interrupted in the middle ( & number & )
                Do
                    off = off + 1
                    number = number + ch
                    If length < off Then Exit Do
                    ch = Mid(biwaz, off, 1)
                    ac = Asc(ch)
                    If ac < 48 Or 58 <= ac Then Exit Do
                Loop
            End Select
        End If

        Value = CDbl(number)
    End Select
End Sub

Public Sub Parse(s, ByRef Value)
    Dim ary, cs, i, j
    ary = Split(s, """")
    ReDim ary2(UBound(ary) / 2)

    ' 制御文字検出第１ステップ
    For i = 0 To 1
        If 0 < InStr(s, Chr(i)) Then Err.Raise 32000, "json parse", "禁則文字chr(" & i & ")が使われています" ' illegal chr ( & i & ) are used
    Next

    ' 文字列配列の抽出
    i = 0
    j = 1
    Do While j <= UBound(ary)
        ary2(i) = ary(j)
        ary(j) = """"
        Do While Right(ary2(i), 1) = "\"
            j = j + 1
            ary2(i) = ary2(i) + """" + ary(j)
            ary(j) = ""
        Loop
        i = i + 1
        j = j + 2
    Loop

    cs = Replace(Join(ary2, Chr(0)), "\\", Chr(1))

    ' 制御文字検出第２ステップ
    For i = 2 To 31
        If 0 < InStr(cs, Chr(i)) Then Err.Raise 32000, "json parse", "禁則文字 chr(" & i & ") が使われています" ' illegal chr ( & i & ) are used
    Next

    ' エスケープ文字の復元
    cs = Replace(Replace(Replace(Replace(Replace(Replace(Replace(cs, "\b", Chr(8)), "\t", vbTab), "\n", vbLf), "\f", vbFormFeed), "\r", vbCr), "\""", """"), "\/", "/")
    ReDim ary2(0)

    i = InStr(cs, "\u")
    If 0 < i Then
        Do
            j = i
            cs = Replace(cs, Mid(cs, j, 6), ChrW("&H" & Mid(cs, j + 2, 4)))
            i = InStr(j + 1, cs, "\u")
        Loop While 0 < i
    End If

    ' 無効なエスケープ文字の検出
    If 0 < InStr(cs, "\") Then Err.Raise 32000, "json parse", "無効なエスケープ '\" & Mid(cs, InStr(cs, "\") + 1, 1) & "' が使われています" ' Invalid escape '\ & Mid (cs, InStr (cs, "\") + 1, 1) & ' is used

    idx = 0
    designs = Split(Replace(cs, Chr(1), "\"), Chr(0))
    off = 1
    biwaz = Replace(Replace(Replace(Replace(Join(ary, ""), vbTab, ""), vbLf, ""), vbCr, ""), " ", "")
    ary = Null

    ParseCore Value
    designs = Null

    If off <= Len(biwaz) Then Err.Raise 32000, "json parse", "json が完結していません ... " & Mid(biwaz, off, 6) ' json is not complete ... & Mid (biwaz, off, 6)
    biwaz = Null
End Sub
