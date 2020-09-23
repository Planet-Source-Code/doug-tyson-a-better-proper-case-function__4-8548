<div align="center">

## A Better Proper Case Function


</div>

### Description

Just copy/paste this function into your code, and it will allow you to convert a string to proper case. Now you've got UCase, LCase, AND PCase.
 
### More Info
 
I tried to account for as many "unimportant" words as I could, but I'm sure I've missed some. Just add any entries you feel necessary in the select statement.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Doug Tyson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/doug-tyson.md)
**Level**          |Intermediate
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/doug-tyson-a-better-proper-case-function__4-8548/archive/master.zip)





### Source Code

```
Function PCase(strInput)
	'Variable declaration.
	dim strArr
	dim tmpWord
	dim tmpString
	dim last
	'Create an array to store each word in the string separately.
	strArr = split(strInput," ")
	if ubound(strArr) > 0 then
		for x = lbound(strArr) to ubound(strArr)
			'Set each word to lower case initially.
			strArr(x) = LCase(strArr(x))
			'Skip the unimportant words.
			select case strArr(x)
				case "a"
				case "an"
				case "and"
				case "but"
				case "by"
				case "for"
				case "in"
				case "into"
				case "is"
				case "of"
				case "off"
				case "on"
				case "onto"
				case "or"
				case "the"
				case "to"
				case "a.m."
					strArr(x) = "A.M."
				case "p.m."
					strArr(x) = "P.M."
				case "b.c."
					strArr(x) = "B.C."
				case "a.d."
					strArr(x) = "A.D."
				case else
					'Capitalize the first letter, but don't forget to take into account that
					'the string may be in single or double quotes.
					if len(strArr(x)) > 1 then
						if mid(strArr(x),1,1) = "'" or mid(strArr(x),1,1) = """" then
							tmpWord = mid(strArr(x),1,1) & Ucase(mid(strArr(x),2,1)) & mid(strArr(x),3,len(strArr(x))-2)
						else
							tmpWord = Ucase(mid(strArr(x),1,1)) & mid(strArr(x),2,len(strArr(x))-1)
						end if
						strArr(x) = tmpWord
					end if
			end select
			'The unimportant words may need to be capitalized if they follow a dash, colon,
			'semi-colon, single quote or double quote.
			if x > 0 then
				if instr(strArr(x-1),"-") _
				or instr(strArr(x-1),":") _
				or instr(strArr(x-1),";") then
					tmpWord = Ucase(mid(strArr(x),1,1)) & mid(strArr(x),2,len(strArr(x))-1)
					strArr(x) = tmpWord
				end if
			end if
		next
	else
		strArr(0) = LCase(strArr(0))
	end if
	'Make sure the first word in the array is upper case, but don't forget to take into account
	'that the string may be in single or double quotes.
	if mid(strArr(0),1,1) = "'" or mid(strArr(0),1,1) = """" then
		tmpWord = mid(strArr(0),1,1) & Ucase(mid(strArr(0),2,1)) & mid(strArr(0),3,len(strArr(0))-2)
	else
		tmpWord = Ucase(mid(strArr(0),1,1)) & mid(strArr(0),2,len(strArr(0))-1)
	end if
	strArr(0) = tmpWord
	'Also, make sure the last word in the array is upper case, but don't forget to take into account
	'that the string may be in single or double quotes.
	last = ubound(strArr)
	if mid(strArr(last),1,1) = "'" or mid(strArr(last),1,1) = """" then
		tmpWord = mid(strArr(last),1,1) & Ucase(mid(strArr(last),2,1)) & mid(strArr(0),3,len(strArr(last))-2)
	else
		tmpWord = Ucase(mid(strArr(last),1,1)) & mid(strArr(last),2,len(strArr(last))-1)
	end if
	strArr(last) = tmpWord
	'Rebuild the whole string from the array parts.
	for x = lbound(strArr) to ubound(strArr)
		tmpString = tmpString & strArr(x) & " "
	next
	PCase = tmpString
End Function
```

