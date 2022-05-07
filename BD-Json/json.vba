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

Public isFast

Private biwaz, design, off, idx, whitespace

Private Function StringifyTab(obj, ByVal off)
    Dim i, tp, sep, ary, key, data
    tp = VarType(obj)
    Select Case tp
    Case vbString
        If IsNull(whitespace) Then
            StringifyTab = """" & obj & """"
        Else
            StringifyTab = """" & Replace(Replace(Replace(Replace(Replace(Replace(Replace(obj, "\", "\\"), Chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), """", "\""") & """"
        End If
    Case vbObject
        Select Case TypeName(obj)
        Case "Dictionary"
            If 0 < obj.Count Then
                ReDim ary(obj.Count - 1)
                i = 0
                If IsNull(whitespace) Then
                    For Each key In obj.Keys
                        ary(i) = StringifyTab(key, off + 1) & ":" & StringifyTab(obj(key), off + 1)
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
            If -1 < UBound(obj) Then
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
        Stringify = Replace(Replace(Replace(Replace(Replace(Replace(Replace(StringifyTab(obj, 0), "\", "\\"), Chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), """", "\""")
    Else
        Stringify = StringifyTab(obj, 0)
    End If
    Dim i
    For Each i In Array(0, 1, 2, 3, 4, 5, 6, 7, 11, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
        Stringify = Replace(Stringify, Chr(i), "\u00" & Right("0" & Hex(i), 2))
    Next
End Function

Private Sub ParseCore(ByRef value)
    Dim ch, child

    ch = Mid(design, off, 1)
    Select Case ch
    Case ""
         off = 1
         idx = idx + 2
         design = biwaz(idx)
         value = biwaz(idx - 1)
    Case "{"
        Set value = CreateObject("Scripting.Dictionary")
        off = off + 1
        ch = Mid(design, off, 1)
        If ch = "}" Then
            off = off + 1
            Exit Sub
        End If

        Do
            If ch <> "" Then Err.Raise 32000, "json parse", "オブジェクトのキーが検出できません" ' Unable to find key of object
            off = 1
            idx = idx + 2
            design = biwaz(idx)
            If Mid(design, off, 1) <> ":" Then Err.Raise 32000, "json parse", "オブジェクトのキー" & design & "の次に ':' を検出できません" ' Unable to find ':' next to object key & design
            off = off + 1
            ch = biwaz(idx - 1)

            ParseCore child
            If VarType(child) = 9 Then Set value(ch) = child Else value(ch) = child

            child = ch
            ch = Mid(design, off, 1)
            If ch = "}" Then
                off = off + 1
                Exit Sub
            ElseIf ch <> "," Then
                Err.Raise 32000, "json parse", "オブジェクトのメンバー """ & child & """:.. の次に ',' を検出できません" ' Unable to find ',' next to object member & child & : ..
            End If
            
            off = off + 1
            ch = Mid(design, off, 1)
        Loop
    Case "["
        Set value = New Collection
        off = off + 1
        ch = Mid(design, off, 1)
        If ch = "]" Then
            off = off + 1
            Exit Sub
        End If

        Do
            ParseCore child
            value.Add child
                
            ch = Mid(design, off, 1)
            If ch = "]" Then
                off = off + 1
                Exit Sub
            ElseIf ch <> "," Then
                Err.Raise 32000, "json parse", "配列要素の次に ',' を検出できません" ' Unable to find ',' next to array element
            End If
            
            off = off + 1
            ch = Mid(design, off, 1)
        Loop
    Case "t"
        If Mid(design, off, 4) <> "true" Then Err.Raise 32000, "json parse", "'t' の次に 'rue' が検出できません" ' 'rue' cannot be detected after 't'
        off = off + 4
        value = True
    Case "f"
        If Mid(design, off, 5) <> "false" Then Err.Raise 32000, "json parse", "'f' の次に 'alse' が検出できません" ' 'alse' cannot be detected after 'f'
        off = off + 5
        value = False
    Case "n"
        If Mid(design, off, 4) <> "null" Then Err.Raise 32000, "json parse", "'n' の次に 'ull' が検出できません" ' 'ull' cannot be detected after 'n'
        off = off + 4
        value = Null
    Case Else
        Dim length, org, ac
        length = Len(design)
        org = off
        If ch = "-" Then
            off = off + 1
            If length < off Then Err.Raise 32000, "json parse", "数値が記号 - の後、途切れています" ' The number is broken after the symbol-
            ch = Mid(design, off, 1)
        End If

        ' integer
        off = off + 1
        ac = Asc(ch)
        If 48 < ac And ac < 58 Then
            Do Until length < off
                ch = Mid(design, off, 1)
                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Exit Do
                off = off + 1
            Loop
        ElseIf ac <> 48 Then
            Err.Raise 32000, "json parse", "不明なトークンです (" & Mid(design, org, off - org) & ")" ' Unknown token ( & mid(design, org, off-org) & )
        End If

        ' fraction
        If off <= length Then
            ch = Mid(design, off, 1)
            If ch = "." Then
                off = off + 1
                If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & Mid(design, org, off - org) & ")" ' The numbers are interrupted in the middle ( & mid(design, org, off-org) & )

                ch = Mid(design, off, 1)
                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & Mid(design, org, off - org) & ")" ' The numbers are interrupted in the middle ( & mid(design, org, off-org) & )

                Do
                    off = off + 1
                    If length < off Then Exit Do
                    ch = Mid(design, off, 1)
                    ac = Asc(ch)
                Loop Until ac < 48 Or 58 <= ac
            End If
        End If

        ' exponent
        If off <= length Then
            Select Case ch
            Case "E", "e"
                off = off + 1
                If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & Mid(design, org, off - org) & ")" ' The numbers are interrupted in the middle ( & mid(design, org, off-org) & )

                ch = Mid(design, off, 1)
                Select Case ch
                Case "-", "+"
                    off = off + 1
                    If length < off Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & Mid(design, org, off - org) & ")" ' The numbers are interrupted in the middle ( & mid(design, org, off-org) & )
                    ch = Mid(design, off, 1)
                End Select

                ac = Asc(ch)
                If ac < 48 Or 58 <= ac Then Err.Raise 32000, "json parse", "数値が途中で途切れています (" & Mid(design, org, off - org) & ")" ' The numbers are interrupted in the middle ( & mid(design, org, off-org) & )
                Do
                    off = off + 1
                    If length < off Then Exit Do
                    ch = Mid(design, off, 1)
                    ac = Asc(ch)
                    If ac < 48 Or 58 <= ac Then Exit Do
                Loop
            End Select
        End If

        value = Val(Mid(design, org, off - org))
    End Select
End Sub

Public Sub Parse(s, ByRef value)
    Dim i, j

If isFast Or IsEmpty(isFast) Then
    design = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, vbCr, ""), vbLf, ""), vbTab, ""), "\\", Chr(0)), "\""", Chr(1)), "\b", Chr(8)), "\t", vbTab), "\n", vbLf), "\f", vbFormFeed), "\r", vbCr), "\/", "/")

    i = InStr(design, "\u")
    If 0 < i Then
        Do
            j = i
            design = Replace(design, Mid(design, j, 6), ChrW("&H" & Mid(design, j + 2, 4)))
            i = InStr(j + 1, design, "\u")
        Loop While 0 < i
    End If

    design = Replace(design, "\", "")

    biwaz = Split(Replace(design, Chr(0), "\"), """")
    For i = 1 To UBound(biwaz) Step 2
        biwaz(i) = Replace(biwaz(i), Chr(1), """")
        biwaz(i + 1) = Replace(biwaz(i + 1), " ", "")
    Next
    If 0 < UBound(biwaz) Then design = biwaz(0)
    design = Replace(design, " ", "")
Else
    For Each i In Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
        If 0 < InStr(s, Chr(i)) Then Err.Raise 32000, "json parse", "禁則文字chr(" & i & ")が使われています" ' illegal chr ( & i & ) are used
    Next
    
    design = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, "\\", Chr(0)), "\""", Chr(1)), "\b", Chr(8)), "\f", vbFormFeed), "\/", "/"), "\r", Chr(2)), "\n", Chr(3)), "\t", Chr(4))

    i = InStr(design, "\u")
    If 0 < i Then
        Do
            j = i
            design = Replace(design, Mid(design, j, 6), ChrW("&H" & Mid(design, j + 2, 4)))
            i = InStr(j + 1, design, "\u")
        Loop While 0 < i
    End If

    If 0 < InStr(design, "\") Then Err.Raise 32000, "json parse", "無効なエスケープ '\" & Mid(design, InStr(design, "\") + 1, 1) & "' が使われています" ' Invalid escape '\ & Mid (design, InStr (design, "\") + 1, 1) & ' is used

    biwaz = Split(Replace(design, Chr(0), "\"), """")
    If 0 < UBound(biwaz) Then
        For i = 0 To UBound(biwaz) Step 2
            biwaz(i) = Replace(Replace(Replace(Replace(biwaz(i), " ", ""), vbCr, ""), vbLf, ""), vbTab, "")
        Next
        design = Join(biwaz, "")
        If 0 < InStr(design, vbTab) Then Err.Raise 32000, "json parse", "文字列中にタブ文字が含まれます" ' detect tab in string
        If 0 < InStr(design, vbCr) Then Err.Raise 32000, "json parse", "文字列中にキャリッジリターン文字が含まれます" ' detect cr in string
        If 0 < InStr(design, vbLf) Then Err.Raise 32000, "json parse", "文字列中にラインフィード文字が含まれます" ' detect lf in string
        For i = 1 To UBound(biwaz) Step 2
            biwaz(i) = Replace(Replace(Replace(Replace(biwaz(i), Chr(1), """"), Chr(2), vbCr), Chr(3), vbLf), Chr(4), vbTab)
        Next
        design = biwaz(0)
    Else
        design = Replace(Replace(Replace(Replace(design, " ", ""), vbCr, ""), vbLf, ""), vbTab, "")
    End If
End If

    idx = 0
    off = 1

    ParseCore value

    If 0 < UBound(biwaz) Then
        If off <= Len(biwaz(idx)) Or idx < UBound(biwaz) Then Err.Raise 32000, "json parse", "json が完結していません ... " ' json is not complete ...
    Else
        If off <= Len(design) Then Err.Raise 32000, "json parse", "json が完結していません ... " ' json is not complete ...
    End If
    biwaz = Null
    design = Null
End Sub
