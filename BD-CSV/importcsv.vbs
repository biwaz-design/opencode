''
' VBS csv parse v1.0.0
' (c) BIWAZ DESIGN - Takeshi Matsui - https://github.com/biwaz-design/opencode/blob/main/BD-CSV/importcsv.vbs
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
option explicit

function readfields(objStream, delim)
	if objStream.atendofstream then
		readfields = null
		exit function
	end if

	dim chunk, quote, pos, s, i

	s = replace(objStream.readline, chr(0), "")
	if instr("," + s, ",""") <= 0 then
		if s = "" then readfields = array("") else readfields = split(s, delim)
		exit function
	end if

	i = 0
	redim record(0)
	chunk = split(s, delim)

	do
		if left(chunk(i), 1) <> """" then
			record(ubound(record)) = chunk(i)
		else
			quote = ""
			s = replace(chunk(i), """""", vbCrLf, 2)
			do
				pos = instr(s, """")
				if 0 < pos then
					if pos <> len(s) then
						quote = quote + replace(left(s, pos - 1), vbCrLf, """") + replace(s, vbCrLf, """""", pos + 1)
					else
						quote = quote + replace(left(s, pos - 1), vbCrLf, """")
					end if
					exit do
				end if

				if i = ubound(chunk) then
					quote = quote + replace(s, vbCrLf, """") + vbCrLf
					if objStream.atendofstream then exit do
					s = replace(objStream.readline, chr(0), "")
					if s = "" then chunk = array("") else chunk = split(s, delim)
					i = -1
				else
					quote = quote + replace(s, vbCrLf, """") + delim
				end if

				i = i + 1
				s = replace(chunk(i), """""", vbCrLf)
			loop
			record(ubound(record)) = quote
		end if

		i = i + 1
		if ubound(chunk) < i then exit do
        redim preserve record(ubound(record) + 1)
	loop

	readfields = record
end function

function writefields(record)
	dim i, result()
	redim result(ubound(record))
	for i = 0 to ubound(record)
		if 0 < instr(record(i), ",") or 0 < instr(record(i), vbCr) or 0 < instr(record(i), vbLf) or left(record(i), 1) = """" then
			result(i) = """" & replace(record(i), """", """""") & """"
		else
			result(i) = record(i)
		end if
	next

	writefields = join(result, ",")
end function

class csv
	private delim, lines_lf, lineno_lf, eof_lf, lines_cr, lineno_cr, eof_cr

	private sub class_initialize()
		lines_lf = array("")
		lines_cr = array("")
		lineno_lf = 1
		lineno_cr = 0
		delim = ","
	end sub

	public sub init(readall, delimiter)
		readall = replace(readall, chr(0), "")
		dim n
		n = len(readall)
		eof_lf = right(readall, 1) = vbLf
		if eof_lf then readall = left(readall, n - 1)

		lines_lf = split(readall, vbLf)

		lineno_lf = 0
		lineno_cr = 0
		delim = delimiter
	end sub

	public function readfields()
		dim chunk, quote, pos, s, i, n

		if lineno_cr = 0 then
			if ubound(lines_lf) < lineno_lf then readfields = null: exit function
			n = len(lines_lf(lineno_lf))
			eof_cr = (right(lines_lf(lineno_lf), 1) = vbCr)
			if eof_cr then lines_lf(lineno_lf) = left(lines_lf(lineno_lf), n - 1): n = n - 1
			if 0 < n then lines_cr = split(lines_lf(lineno_lf), vbCr) else lines_cr = array("")
			lineno_lf = lineno_lf + 1
		end if

		s = lines_cr(lineno_cr)
		if lineno_cr < ubound(lines_cr) then lineno_cr = lineno_cr + 1 else lineno_cr = 0

		if instr("," + s, ",""") <= 0 then
			if s = "" then readfields = array("") else readfields = split(s, delim)
			exit function
		end if

		i = 0
		redim record(0)
		chunk = split(s, delim)

		do
			if left(chunk(i), 1) <> """" then
				record(ubound(record)) = chunk(i)
			else
				quote = ""
				s = replace(chunk(i), """""", vbCrLf, 2)
				do
					pos = instr(s, """")
					if 0 < pos then
						if pos <> len(s) then
							quote = quote + replace(left(s, pos - 1), vbCrLf, """") + replace(s, vbCrLf, """""", pos + 1)
						else
							quote = quote + replace(left(s, pos - 1), vbCrLf, """")
						end if
						exit do
					end if

					if i = ubound(chunk) then
						if lineno_cr = 0 then
							if eof_cr then eof_cr = vbCrLf else eof_cr = vbLf
							quote = quote + replace(s, vbCrLf, """") + eof_cr
							if ubound(lines_lf) < lineno_lf then
								if not eof_lf then quote = left(quote, len(quote) - 1)
								exit do
							end if
							n = len(lines_lf(lineno_lf))
							eof_cr = (right(lines_lf(lineno_lf), 1) = vbCr)
							if eof_cr then lines_lf(lineno_lf) = left(lines_lf(lineno_lf), n - 1): n = n - 1
							if 0 < n then lines_cr = split(lines_lf(lineno_lf), vbCr) else lines_cr = array("")
							lineno_lf = lineno_lf + 1
						else
							quote = quote + replace(s, vbCrLf, """") + vbCr
						end if

						s = lines_cr(lineno_cr)
						if lineno_cr < ubound(lines_cr) then lineno_cr = lineno_cr + 1 else lineno_cr = 0

						if s = "" then chunk = array("") else chunk = split(s, delim)
						i = -1
					else
						quote = quote + replace(s, vbCrLf, """") + delim
					end if

					i = i + 1
					s = replace(chunk(i), """""", vbCrLf)
				loop
				record(ubound(record)) = quote
			end if

			i = i + 1
			if ubound(chunk) < i then exit do
        	redim preserve record(ubound(record) + 1)
		loop

		readfields = record
	end function
end class

sub SelfSub
	dim i, j, n, isUtf8, isQuote, objCsv, objStream, record, result(), start

	if WScript.Arguments.Count = 0 then
		wscript.echo "usage : cscript //nologo importcsv.vbs /u /100 [target.txt] > [output.csv]"
		wscript.echo "		/q   ... without quote process"
		wscript.echo "		/u   ... Read As Utf-8(default is Shift-JIS)"
		wscript.echo "		/100 ... Parse times(only when parse time less than 2, parsed result will write out)"
	end if

	n = 1
	isUtf8 = false
	isQuote = true
	for i = 0 to wscript.arguments.count - 1
		if left(wscript.arguments(i), 1) = "/" then
			select case mid(WScript.Arguments(i), 2, 1)
			case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
				n = cint(mid(WScript.Arguments(i), 2))
			case "u"
				isUtf8 = not isUtf8
			case "q"
				isQuote = not isQuote
			end select
		else
			if n < 2 then
				if isUtf8 then
					set objCsv = new csv
					with createobject("ADODB.Stream")
						.charset = "utf-8"
						.type = 1
						.open
						.loadfromfile WScript.Arguments(i)
						.position = 0
						.type = 2
						objCsv.init .readtext, ","
						.close
					end with

					redim result(-1)
					do
						record = objCsv.readfields()
						if isnull(record) then exit do
						redim preserve result(ubound(result) + 1)
						result(ubound(result)) = writefields(record)
					loop
					wscript.stdout.write join(result, vbCrLf)
					set objCsv = nothing
				else
					set objStream = createobject("Scripting.FileSystemObject").opentextfile(wscript.arguments(i))
					redim result(-1)
					do
						record = readfields(objStream, ",")
						if isnull(record) then exit do
						redim preserve result(ubound(result) + 1)
						result(ubound(result)) = writefields(record)
					loop
					wscript.stdout.write join(result, vbCrLf)
					objStream.close
					set objStream = nothing
				end if
			else
				start = now

				for j = 1 to n
					if isUtf8 then
						dim readall
						with createobject("ADODB.Stream")
							.charset = "utf-8"
							.type = 1
							.open
							.loadfromfile WScript.Arguments(i)
							.position = 0
							.type = 2
							readall = .readtext
							.close
						end with

						if isQuote then
							set objCsv = new csv
							objCsv.init readall, ","

							do
								record = objCsv.readfields()
								if isnull(record) then exit do
							loop
							set objCsv = nothing
						else
							dim s, lines_lf, lineno_lf, lines_cr, lineno_cr

							n = len(readall)
							if (right(readall, 1) = vbLf) then readall = left(readall, n - 1)
							lines_lf = split(readall, vbLf)
							lineno_lf = 0
							lineno_cr = 0

							do
								if lineno_cr = 0 then
									if ubound(lines_lf) < lineno_lf then exit do
									n = len(lines_lf(lineno_lf))
									if (right(lines_lf(lineno_lf), 1) = vbCr) then lines_lf(lineno_lf) = left(lines_lf(lineno_lf), n - 1) : n = n - 1
									if 0 < n then lines_cr = split(lines_lf(lineno_lf), vbCr) else lines_cr = array("")
									lineno_lf = lineno_lf + 1
								end if

								s = lines_cr(lineno_cr)
								if lineno_cr < ubound(lines_cr) then lineno_cr = lineno_cr + 1 else lineno_cr = 0

								if s = "" then exit do
								split s, ","
							loop
						end if
					else
						set objStream = createobject("Scripting.FileSystemObject").opentextfile(wscript.arguments(i))
						if isQuote then
							do
								record = readfields(objStream, ",")
								if isnull(record) then exit do
							loop
						else
							do until objStream.atendofstream
								record = split(objStream.readline, ",")
							loop
						end if
						objStream.close
						set objStream = nothing
					end if
				next

				WScript.StdErr.WriteLine datediff("s", start, now)
			end if
		end if
	next
end sub

if WScript.ScriptName = "importcsv.vbs" then SelfSub
