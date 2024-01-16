''
' VBA csv parse v1.0.0
' (c) BIWAZ DESIGN - Takeshi Matsui - https://github.com/biwaz-design/opencode/blob/main/BD-CSV/importcsv.bas
'
' This is a completely original CSV parser, designed with (probably)
' a unique parsing method, as it would take too much time to parse each
' character individually. I hope it will be useful for your work.
' -----------------------------------------------------------------------
' 完全にオリジナルのCSVパーサーです。１文字単位で解析を行うと時間が掛かって
' 仕方がないですので、（恐らく）独自の解析方法にて設計しました。
' お役に立てば幸いです。※SHIFT-JISで保存しなおして、ご利用ください。
'
' * Parse textfile stream to object
' * Parse quoted csv-string to object
' * Stringify object(array or collection) to quoted csv-string
'
' @class CSV Converter
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

Function readfields(objStream, Optional delim As String = ",")
    If objStream.atendofstream Then
        readfields = Null
        Exit Function
    End If

    Dim chunk, quote, pos, s, i

    s = Replace(objStream.readline, Chr(0), "")
    If InStr("," + s, ",""") <= 0 Then
        If s = "" Then readfields = Array("") Else readfields = Split(s, delim)
        Exit Function
    End If

    i = 0
    ReDim record(0)
    chunk = Split(s, delim)

    Do
        If Left(chunk(i), 1) <> """" Then
            record(UBound(record)) = chunk(i)
        Else
            quote = ""
            s = Replace(chunk(i), """""", vbCrLf, 2)
            Do
                pos = InStr(s, """")
                If 0 < pos Then
                    If pos <> Len(s) Then
                        quote = quote + Replace(Left(s, pos - 1), vbCrLf, """") + Replace(s, vbCrLf, """""", pos + 1)
                    Else
                        quote = quote + Replace(Left(s, pos - 1), vbCrLf, """")
                    End If
                    Exit Do
                End If

                If i = UBound(chunk) Then
                    quote = quote + Replace(s, vbCrLf, """") + vbCrLf
                    If objStream.atendofstream Then Exit Do
                    s = Replace(objStream.readline, Chr(0), "")
                    If s = "" Then chunk = Array("") Else chunk = Split(s, delim)
                    i = -1
                Else
                    quote = quote + Replace(s, vbCrLf, """") + delim
                End If

                i = i + 1
                s = Replace(chunk(i), """""", vbCrLf)
            Loop
            record(UBound(record)) = quote
        End If

        i = i + 1
        If UBound(chunk) < i Then Exit Do
        ReDim Preserve record(UBound(record) + 1)
    Loop

    readfields = record
End Function

Function writefields(record)
    Dim i, cell, result()

    Select Case TypeName(record)
    Case "Collection"
        If record.Count = 0 Then writefields = "": Exit Function
        ReDim result(record.Count - 1)
        i = 0
        For Each cell In record
            If 0 < InStr(cell, ",") Or 0 < InStr(cell, vbCr) Or 0 < InStr(cell, vbLf) Or Left(cell, 1) = """" Then
                result(i) = """" & Replace(cell, """", """""") & """"
            Else
                result(i) = cell
            End If
            i = i + 1
        Next
    Case Else
        If UBound(record) = -1 Then writefields = "": Exit Function
        ReDim result(UBound(record))
        For i = 0 To UBound(record)
            If 0 < InStr(record(i), ",") Or 0 < InStr(record(i), vbCr) Or 0 < InStr(record(i), vbLf) Or Left(record(i), 1) = """" Then
                result(i) = """" & Replace(record(i), """", """""") & """"
            Else
                result(i) = record(i)
            End If
        Next
    End Select

    writefields = Join(result, ",")
End Function