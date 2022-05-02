''
' VBS json parse v1.0.0
' (c) BIWAZ DESIGN - Takeshi Matsui - https://github.com/biwaz-design/opencode
'
' It's a completely original JSON parser, but I think the structure will not 
' change much no matter who makes it. I wanted to improve the speed of the 
' off-the-shelf parser as much as possible, so I devised it. I hope it will be 
' useful for your work.
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
' Redistribution and use in biwaz and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of biwaz code must retain the above copyright
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
option explicit

private biwaz, designs, off, idx, whitespace

private function StringifyTab(obj, byval off)
	dim i, tp, sep, ary, key, data
	tp = vartype(obj)
	select case tp
	case 8
		if isnull(whitespace) then
			StringifyTab = """" & obj & """"
		else
			StringifyTab = """" & replace(replace(replace(replace(replace(replace(replace(obj, "\", "\\"), chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), "/", "\/") & """"
		end if
	case 9
		if 0 < obj.count then
			redim ary(obj.count - 1)
			i = 0
			if isnull(whitespace) then
				for each key in obj.keys
					ary(i) = sep + StringifyTab(key, off + 1) & ":" & StringifyTab(obj(key), off + 1)
					i = i + 1
				next
				ary(ubound(ary)) = ary(ubound(ary)) & "}"
			else
				sep = vbCrLf & string(off + 1, whitespace)
				for each key in obj.keys
					ary(i) = sep + StringifyTab(key, off + 1) & ": " & StringifyTab(obj(key), off + 1)
					i = i + 1
				next
				ary(ubound(ary)) = ary(ubound(ary)) & left(sep, len(sep) - 1) & "}"
			end if
			ary(0) = "{" & ary(0)
			StringifyTab = join(ary, ",")
		else
			StringifyTab = "{}"
		end if
	case 1
		StringifyTab = "null"
	case 11
		if obj then StringifyTab = "true" else StringifyTab = "false"
	case 7
		StringifyTab = """" & obj & """"
	case 17
		StringifyTab = """--binary data--"""
	case else
		if 8192 <= tp then
			if -1 < ubound(obj) then
				redim ary(ubound(obj))
				i = 0
				if isnull(whitespace) then
					for each data in obj
						ary(i) = StringifyTab(data, off + 1)
						i = i + 1
					next
					ary(ubound(ary)) = ary(ubound(ary)) & "]"
				else
					sep = vbCrLf & string(off + 1, whitespace)
					for each data in obj
						ary(i) = sep & StringifyTab(data, off + 1)
						i = i + 1
					next
					ary(ubound(ary)) = ary(ubound(ary)) & left(sep, len(sep) - 1) & "]"
				end if
				ary(0) = "[" & ary(0)
				StringifyTab = join(ary, ",")
			else
				StringifyTab = "[]"
			end if
		else
			StringifyTab = obj
		end if
	end select
end function

public function Stringify(obj, ws)
	whitespace = ws
	if isnull(whitespace) then
		Stringify = replace(replace(replace(replace(replace(replace(replace(StringifyTab(obj, 0), "\", "\\"), chr(8), "\b"), vbTab, "\t"), vbLf, "\n"), vbFormFeed, "\f"), vbCr, "\r"), "/", "\/")
	else
		Stringify = StringifyTab(obj, 0)
	end if
	dim i
	for each i in array(0, 1, 2, 3, 4, 5, 6, 7, 11, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
		Stringify = replace(Stringify, chr(i), "\u00" & right("0" & hex(i), 2))
	next
end function

private sub ParseCore(byref value)
	dim ch, child

	ch = mid(biwaz, off, 1)
	select case ch
	case "{"
		set value = createobject("Scripting.Dictionary")
		off = off + 1
		ch = mid(biwaz, off, 1)
		if ch = "}" then
			off = off + 1
			exit sub
		end if

		do
			if ch <> """" then err.raise 32000, "json parse", "オブジェクトのキーが検出できません"
			off = off + 1
			if mid(biwaz, off, 1) <> ":" then err.raise 32000, "json parse", "オブジェクトのキー" & designs(idx) & "の次に ':' を検出できません"
			ch = designs(idx)
			idx = idx + 1

			off = off + 1
			ParseCore child
			if vartype(child) = 9 then set value(ch) = child else value(ch) = child

			child = ch
			ch = mid(biwaz, off, 1)
			if ch = "}" then 
				off = off + 1
				exit sub
			elseif ch <> "," then
				err.raise 32000, "json parse", "オブジェクトのメンバー """ & child & """:.. の次に ',' を検出できません"
			end if

			off = off + 1
			ch = mid(biwaz, off, 1)
		loop
	case "["
		redim value(-1)
		off = off + 1
		ch = mid(biwaz, off, 1)
		if ch = "]" then
			off = off + 1
			exit sub
		end if

		do
			ParseCore child
			redim preserve value(ubound(value) + 1)
			if vartype(child) = 9 then set value(ubound(value)) = child else value(ubound(value)) = child

			ch = mid(biwaz, off, 1)
			if ch = "]" then 
				off = off + 1
				exit sub
			elseif ch <> "," then
				err.raise 32000, "json parse", "配列要素の次に ',' を検出できません"
			end if
			off = off + 1
		loop
	case """"
		off = off + 1
		value = designs(idx)
		idx = idx + 1
	case "t"
		if mid(biwaz, off, 4) <> "true" then err.raise 32000, "json parse", "'t' の次に 'rue' が検出できません"
		off = off + 4
		value = true
	case "f"
		if mid(biwaz, off, 5) <> "false" then err.raise 32000, "json parse", "'f' の次に 'alse' が検出できません"
		off = off + 5
		value = false
	case "n"
		if mid(biwaz, off, 4) <> "null" then err.raise 32000, "json parse", "'n' の次に 'ull' が検出できません"
		off = off + 4
		value = null
	case else
		dim length, number, ac
		length = len(biwaz)
		if ch = "-" then
			off = off + 1
			if length < off then err.raise 32000, "json parse", "数値が記号 - の後、途切れています"
			number = ch
			ch = mid(biwaz, off, 1)
		else
			number = ""
		end if

		' integer
		off = off + 1
		number = number + ch
		ac = Asc(ch)
		if 48 < ac and ac < 58 then
			do until length < off
				ch = mid(biwaz, off, 1)
				ac = asc(ch)
				if ac < 48 or 58 <= ac then exit do
				off = off + 1
				number = number + ch
			loop
		elseif ac <> 48 then
			err.raise 32000, "json parse", "不明なトークンです (" & number & ")"
		end if

		' fraction
		if off <= length then
			ch = mid(biwaz, off, 1)
			if ch = "." then
				off = off + 1
				number = number + ch
				if length < off then err.raise 32000, "json parse", "数値が途中で途切れています (" & number & ")"

				ch = mid(biwaz, off, 1)
				ac = asc(ch)
				if ac < 48 or 58 <= ac then err.raise 32000, "json parse", "数値が途中で途切れています (" & number & ")"

				do
					off = off + 1
					number = number + ch
					if length < off then exit do
					ch = mid(biwaz, off, 1)
					ac = asc(ch)
				loop until ac < 48 or 58 <= ac
			end if
		end if

		' exponent
		if off <= length then
			select case ch
			case "E", "e"
				off = off + 1
				number = number + ch
				if length < off then err.raise 32000, "json parse", "数値が途中で途切れています (" & number & ")"

				ch = mid(biwaz, off, 1)
				select case ch
				case "-", "+"
					off = off + 1
					number = number + ch
					if length < off then err.raise 32000, "json parse", "数値が途中で途切れています (" & number & ")"
					ch = mid(biwaz, off, 1)
				end select

				ac = asc(ch)
				if ac < 48 or 58 <= ac then err.raise 32000, "json parse", "数値が途中で途切れています (" & number & ")"
				do
					off = off + 1
					number = number + ch
					if length < off then exit do
					ch = mid(biwaz, off, 1)
					ac = asc(ch)
					if ac < 48 or 58 <= ac then exit do
				loop
			end select
		end if

		value = cdbl(number)
	end select
end sub

public sub Parse(s, byref value)
	dim ary, cs, i, j
	ary = split(s, """")
	redim ary2(ubound(ary) / 2)

	' 制御文字検出第１ステップ
	for i = 0 to 1
		if 0 < instr(s, chr(i)) then err.raise 32000, "json parse", "禁則文字chr(" & i & ")が使われています"
	next

	' 文字列配列の抽出
	i = 0
	j = 1
	do while j <= ubound(ary)
		ary2(i) = ary(j)
		ary(j) = """"
		do while right(ary2(i), 1) = "\"
			j = j + 1
			ary2(i) = ary2(i) + """" + ary(j)
			ary(j) = ""
		loop
		i = i + 1
		j = j + 2
	loop

	cs = replace(join(ary2, chr(0)), "\\", chr(1))

	' 制御文字検出第２ステップ
	for i = 2 to 31
		if 0 < instr(cs, chr(i)) then err.raise 32000, "json parse", "禁則文字 chr(" & i & ") が使われています"
	next

	' エスケープ文字の復元
	cs = replace(replace(replace(replace(replace(replace(replace(cs, "\b", chr(8)), "\t", vbTab), "\n", vbLf), "\f", vbFormFeed), "\r", vbCr), "\""", """"), "\/", "/")
	redim ary2(-1)

	i = instr(cs, "\u")
	if 0 < i then
		do
			j = i
			cs = replace(cs, mid(cs, j, 6), chrw("&H" & mid(cs, j + 2, 4)))
			i = instr(j + 1, cs, "\u")
		loop while 0 < i
	end if

	' 無効なエスケープ文字の検出
	if 0 < instr(cs, "\") then err.raise 32000, "json parse", "無効なエスケープ '\" & mid(cs, instr(cs, "\")+1, 1) & "' が使われています"

	idx = 0
	designs = split(replace(cs, chr(1), "\"), chr(0))
	off = 1
	biwaz = replace(replace(replace(replace(join(ary, ""), vbTab, ""), vbLf, ""), vbCr, ""), " ", "")
	ary = null

	ParseCore value
	designs = null

	if off <= len(biwaz) then err.raise 32000, "json parse", "json が完結していません ... " & mid(biwaz, off, 6)
	biwaz = null
end sub

sub SelfSub
	dim i, s, n, start, value, raptime

	if 0 < WScript.Arguments.Count then
		with createobject("Scripting.FileSystemObject").opentextfile(WScript.Arguments(0))
			s = .readall
			.close
		end with

		if 1 < WScript.Arguments.Count then n = WScript.Arguments(1) else n = 0
		start = now
		select case n
		case 0
			Parse s, value
			raptime = datediff("s", start, now)

			start = now
			wscript.echo Stringify(value, vbTab)
		case else
			for i = 1 to n
				Parse s, value
			next
			raptime = datediff("s", start, now)

			start = now
			if n = 1 then wscript.echo Stringify(value, null)
		end select

		WScript.StdErr.WriteLine raptime & " " & datediff("s", start, now)
	else
		wscript.echo "usage : cscript //nologo json.vbs [target.txt] [parse times]"
	end if
end sub

if WScript.ScriptName = "json.vbs" then SelfSub
