' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const BYTE_OFFSET As Byte = 2

Option Explicit
Public Function ParseJsonString(ByRef json$) As Object
    
    Dim jsonBytes() As Byte
    
    jsonBytes = json
    PreProcessData jsonBytes
    Set ParseJsonString = ParseJSON(jsonBytes)
    
End Function
Public Function ParseJsonBytes(ByRef jsonBytes() As Byte) As Object
    
    PreProcessData jsonBytes
    Set ParseJsonBytes = ParseJSON(jsonBytes)

    'if ParseJsonBytes is Nothing then
    '    err.raise INVALID_JSON_FORMAT_ERROR,"ParseJsonBytes","Bytes couldn't be parsed."
    'End If

End Function

Private Function ParseJSON(ByRef json() As Byte, Optional ByRef currentTokenIndex As Long = 1, Optional maxIndex& = -1, Optional ByVal recursionDepth& = -1, Optional ByVal arrayOrObjectName$ = vbNullString) As Object
''====================================================================================================================================================================================
'' Summary: Parse Json strings.
'' Inputs:
''   json: String to parse.
''   currentTokenIndex: Index to start parsing from. Initialize at 1 or ensure it is a json token.
''   recursionDepth: Used for generating error messages.
''   arrayOrObjectName: Used for generating error messages.
'' Reference:
''   https://www.json.org/json-en.html
''====================================================================================================================================================================================
    Dim keyAvailable As Boolean, value$, unIdentifiedCount&, parsingJsonArray As Boolean, parsingJsonObject As Boolean

    Dim keyForValue$, jsonContainer As Object, keyValueDelimiterTokenIndex&, nextTokenDelimiterIndex&, keyStart&, keyEnd&, valueStart&, valueEnd&, _
        valueAvailable As Boolean, isValueContainedWithinUnescapedQuotes As Boolean
    
    Const quote As Byte = 34, backslash As Byte = 92, rightCurly As Byte = 125, leftCurly As Byte = 123, colon$ = ":", leftSquare As Byte = 91, rightSquare As Byte = 93 ', arrayPattern$ = "[[]*[]]", objectPattern$ = "{*}", quoteEnclosedPattern$ = """*"""

    recursionDepth = recursionDepth + 1: unIdentifiedCount = -1
'
    If recursionDepth = 0 Then
        maxIndex = UBound(json)
        currentTokenIndex = LBound(json)
    End If

    On Error GoTo CATCH_FAILURE
    
    parsingJsonObject = (json(currentTokenIndex) = leftCurly And json(maxIndex - (BYTE_OFFSET - 1)) = rightCurly)
    parsingJsonArray = (json(currentTokenIndex) = leftSquare And json(maxIndex - (BYTE_OFFSET - 1)) = rightSquare)
    
    If parsingJsonObject Or parsingJsonArray Then
        'Max index - 1 should be the last non-null character in the range.
        Do While nextTokenDelimiterIndex < maxIndex - (BYTE_OFFSET - 1)

            nextTokenDelimiterIndex = GetNextTokenDelimiterIndex(json, currentTokenIndex, maxIndex - (BYTE_OFFSET - 1))

            If parsingJsonObject Then
                'Ex: {"apple":45}
                ' If non whitespace characters exist then record their location; otherwise throw an error.
                If TryFindFirstAndLastNonWhitespace(json, currentTokenIndex + BYTE_OFFSET, nextTokenDelimiterIndex - BYTE_OFFSET, keyStart, valueEnd) Then
                    ' Does key start with a quote?
                    If json(keyStart) = quote Then
                        keyValueDelimiterTokenIndex = LocationOfCharacterNotBetweenTextTokens(json, colon, keyStart, valueEnd)
                        ' Is there a colon outside of quotes?
                        If keyValueDelimiterTokenIndex > 0 Then
                            ' Find boundaries of key using location of key-value delimiter.
                            Call TryFindFirstAndLastNonWhitespace(json, keyStart, keyValueDelimiterTokenIndex - BYTE_OFFSET, keyStart, keyEnd)
                            ' Find boundaries of value.
                            valueAvailable = TryFindFirstAndLastNonWhitespace(json, keyValueDelimiterTokenIndex + BYTE_OFFSET, nextTokenDelimiterIndex - BYTE_OFFSET, valueStart, valueEnd)
                            
                            If (keyEnd < valueStart) Then
                                keyEnd = keyEnd + (BYTE_OFFSET - 1)
                                ' Will raise error if invalid key.
                                keyForValue = ProcessSpecialCharactersInJsonStringValue(json, keyStart, keyEnd, isValueContainedWithinUnescapedQuotes)
                                keyAvailable = isValueContainedWithinUnescapedQuotes
                                If Not keyAvailable Then Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Key must be enclosed in quotes." & CStr(GetArraySubset(json, keyStart, keyEnd))
                            Else
                                Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Key overlaps with value." & CStr(GetArraySubset(json, keyStart, valueEnd + (BYTE_OFFSET - 1)))
                            End If
                        Else
                            Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Key-Value delimiter not found." & CStr(GetArraySubset(json, keyStart, valueEnd + (BYTE_OFFSET - 1)))
                        End If
                    Else
                        Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Key doesn't start with a quote: " & CStr(GetArraySubset(json, keyStart, valueEnd + (BYTE_OFFSET - 1)))
                    End If
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Expected Key-Value pair not found."
                End If
            End If

            If parsingJsonArray Or (keyAvailable And parsingJsonObject) Then

                If jsonContainer Is Nothing Then
                    If parsingJsonObject Then
                        #If Mac Then
                            Set jsonContainer = New Dictionary
                        #Else
                            Set jsonContainer = CreateObject("Scripting.Dictionary")
                        #End If
                    Else
                        Set jsonContainer = New Collection
                    End If
                End If
                            
                If parsingJsonArray Then
                    unIdentifiedCount = unIdentifiedCount + 1
                    'Adjust by BYTE_OFFSET to exclude token delimiters.
                    valueAvailable = TryFindFirstAndLastNonWhitespace(json, currentTokenIndex + BYTE_OFFSET, nextTokenDelimiterIndex - BYTE_OFFSET, valueStart, valueEnd)
                End If
                ' Ends are currently set to the first byte in a 2 byte character. Set them to the second to encapsulate.
                valueEnd = valueEnd + (BYTE_OFFSET - 1)

                If valueAvailable And ((json(valueStart) = leftCurly And json(valueEnd - (BYTE_OFFSET - 1)) = rightCurly) Or (json(valueStart) = leftSquare And json(valueEnd - (BYTE_OFFSET - 1)) = rightSquare)) Then
                    'If enclosed in {} or [] then parse a new object or array.
                    If parsingJsonObject Then
                        jsonContainer.Add IIf(keyAvailable, keyForValue, unIdentifiedCount), ParseJSON(json, valueStart, valueEnd, recursionDepth, keyForValue)
                    Else
                        jsonContainer.Add ParseJSON(json, valueStart, valueEnd, recursionDepth, IIf(keyAvailable, keyForValue, unIdentifiedCount))
                    End If
                Else
                    If valueAvailable Then
                        ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                        ' Throws an error if any other unescaped quotes.
                        value = ProcessSpecialCharactersInJsonStringValue(json, valueStart, valueEnd, isValueContainedWithinUnescapedQuotes)
                    Else
                        value = "null"
                    End If
                    
                    If isValueContainedWithinUnescapedQuotes Then
                        ' String value.
                        If parsingJsonObject Then
                            jsonContainer.Add keyForValue, value
                        Else
                            jsonContainer.Add value
                        End If
                    ElseIf IsNumeric(value) Then
                        If parsingJsonObject Then
                            jsonContainer.Add keyForValue, IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
                        Else
                            jsonContainer.Add IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
                        End If
                    Else
                        Select Case value
                            Case "true", "false"
                                If parsingJsonObject Then
                                    jsonContainer.Add keyForValue, CBool(value)
                                Else
                                    jsonContainer.Add CBool(value)
                                End If
                            Case "null", vbNullString
                                If parsingJsonObject Then
                                    jsonContainer.Add keyForValue, Null
                                Else
                                    jsonContainer.Add Null
                                End If
                            Case Else
                                Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid format: Value isn't numeric and failed string literal checks. Ensure it is nclosed in an unescaped quotes: " & value
                        End Select
                    End If
                End If

                If keyAvailable Then
                    keyAvailable = False: keyForValue = vbNullString
                End If
                value = vbNullString
            End If
            currentTokenIndex = nextTokenDelimiterIndex
        Loop

    End If
    Set ParseJSON = jsonContainer
    Exit Function
CATCH_FAILURE:
    'Stop: Resume
    PropagateError Err, "ParseJSON", "Name: '" & arrayOrObjectName & "' Depth: " & recursionDepth & " Item: " & IIf(keyAvailable, keyForValue, unIdentifiedCount)
End Function
Private Function GetNextTokenDelimiterIndex(json() As Byte, currentTokenIndex&, maxTokenIndex&) As Long
'===========================================================================================================
' Summary:
'   Returns the location of the next token that encapsulates a complete value based on [currentTokenIndex]
' Notes:
'   [currentTokenIndex] should be set to one of  ,[{ which will note that the text that follows is a value/array/object.
'===========================================================================================================
    Dim nextSquareClose&, nextSquareOpen&, nextValidCommaTokenIndexBaseOne&, tempIndexBaseOne&, nextOpen&, nextClose&
    
    Dim nextCurlyClose&, nextCurlyOpen&, previousClose&, tokenIndicator$(1), escapeEnclosingArrayOrObject As Boolean, nextTokenDelimiterIndex&
    
    Dim lboundMinusOne&: lboundMinusOne = LBound(json) - 1
    ' These are tokens that denote the start of a new value, or the start/end of a new object or array.
    Const comma$ = ",", leftSquare$ = "[", rightSquare$ = "]", leftCurly$ = "{", rightCurly$ = "}"
    
    ' Set to index of current token in array. Adjusted to be 1 based.
    tempIndexBaseOne = currentTokenIndex - lboundMinusOne
        
    nextCurlyOpen = InStrB(tempIndexBaseOne + 1, json, leftCurly)
    nextCurlyClose = InStrB(tempIndexBaseOne + 1, json, rightCurly)
    nextSquareOpen = InStrB(tempIndexBaseOne + 1, json, leftSquare)
    nextSquareClose = InStrB(tempIndexBaseOne + 1, json, rightSquare)
    
    nextValidCommaTokenIndexBaseOne = LocationOfCharacterNotBetweenTextTokens(json, comma, currentTokenIndex + BYTE_OFFSET, maxTokenIndex) - lboundMinusOne

    If nextValidCommaTokenIndexBaseOne > 0 Then
        
        '===========================================================================================================================
        ' Determine which of [ or { is greater than currentTokenIndex AND appears first.
        ' If nextValidCommaTokenIndexBaseOne is greater than that array or object token start,
        ' then determine if the comma is encapsulated. If it is then find the next comma that isn't encapsulated.
        '===========================================================================================================================
        If nextCurlyOpen < nextValidCommaTokenIndexBaseOne And nextCurlyOpen > 0 And (nextCurlyOpen < nextSquareOpen Or nextSquareOpen = 0) Then
            
            tokenIndicator(0) = leftCurly: tokenIndicator(1) = rightCurly
            nextOpen = nextCurlyOpen: nextClose = nextCurlyClose
            escapeEnclosingArrayOrObject = True
            
        ElseIf nextSquareOpen < nextValidCommaTokenIndexBaseOne And nextSquareOpen > 0 And (nextSquareOpen < nextCurlyOpen Or nextCurlyOpen = 0) Then
            
            tokenIndicator(0) = leftSquare: tokenIndicator(1) = rightSquare
            nextOpen = nextSquareOpen: nextClose = nextSquareClose
            escapeEnclosingArrayOrObject = True
            
        End If
    
        If escapeEnclosingArrayOrObject Then
            'Logic: Initialize previousClose to the next token End character found after the current position in the byte array.
            'Update the next token end and open position.
            'If open greater than the previous close or next open = 0 then you have exited the object/array that triggered the loop.
            Do
                previousClose = nextClose
                nextClose = InStrB(previousClose + BYTE_OFFSET, json, tokenIndicator(1))
                nextOpen = InStrB(nextOpen + BYTE_OFFSET, json, tokenIndicator(0))
            Loop Until nextOpen > previousClose Or nextOpen = 0 Or (previousClose + lboundMinusOne >= maxTokenIndex)
            
            ' Set tempIndexBaseOne to the next character after the end of the array or object that triggered loop.
            tempIndexBaseOne = previousClose + BYTE_OFFSET
            ' Check if still within bounds.
            If tempIndexBaseOne + lboundMinusOne < maxTokenIndex Then
                nextValidCommaTokenIndexBaseOne = LocationOfCharacterNotBetweenTextTokens(json, comma, tempIndexBaseOne + lboundMinusOne, maxTokenIndex)
                If nextValidCommaTokenIndexBaseOne > 0 Then nextValidCommaTokenIndexBaseOne = nextValidCommaTokenIndexBaseOne - lboundMinusOne
            Else
                nextValidCommaTokenIndexBaseOne = 0
            End If
        End If
        '===========================================================================================================================
    End If
    ' Get token index in the same base as input array.
    nextTokenDelimiterIndex = nextValidCommaTokenIndexBaseOne + lboundMinusOne
    
    Select Case nextTokenDelimiterIndex
        Case Is <= 0, currentTokenIndex, Is > maxTokenIndex
            nextTokenDelimiterIndex = maxTokenIndex
    End Select
    
    Select Case json(nextTokenDelimiterIndex)
        Case 44, 125, 93
            ',}] Token is valid.
        Case Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, "GetNextTokenDelimiterIndex", "Invalid JSON format. Calculated token index doesn't reference a value delimiter: " & Chr(json(nextTokenDelimiterIndex))
    End Select
    
    GetNextTokenDelimiterIndex = nextTokenDelimiterIndex
    Exit Function
Propagate:
    PropagateError Err, "GetNextTokenDelimiterIndex"
End Function
Private Function ProcessSpecialCharactersInJsonStringValue(ByRef json() As Byte, ByRef startTokenIndex&, ByRef maxTokenIndex&, Optional ByRef isValueContainedWithinUnescapedQuotes As Boolean) As Byte()
'========================================================================================================================================
'Summary: Converts special characters that are lead by a \.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes backslash
'========================================================================================================================================
    Dim outputBytes() As Byte, iMainLoop&, outputIndex&, ignoreSpecialCharacter As Boolean, lboundMinusOne&, _
    groupedBackslashCount&, countOfBackslashToKeep&, tempCount&, uniCode(8) As Byte, processBytes As Boolean, specialChars(6) As String, unescapedQuoteCount As Long
    
    Const quote As Byte = 34, backslash As Byte = 92

    specialChars(0) = Chr(quote): specialChars(1) = Chr(backslash)
    specialChars(2) = vbCr: specialChars(3) = vbTab
    specialChars(4) = vbLf: specialChars(5) = VBA.vbBack
    specialChars(6) = VBA.vbFormFeed

    '(LBound(json) - 1) will adjust for non 1 based arrays.
    lboundMinusOne = LBound(json) - 1
    ' Determine if there are any special characters in the byte range.
    For iMainLoop = LBound(specialChars) To UBound(specialChars)
        tempCount = InStrB(startTokenIndex - lboundMinusOne, json, specialChars(iMainLoop))
        ' If within bounds set processBytes equal to true and exit the for loop.
        If tempCount <= (maxTokenIndex - lboundMinusOne) And tempCount > 0 Then
            processBytes = True: Exit For
        End If
    Next iMainLoop

    On Error GoTo PropagateError
    
    If processBytes Then
    
        ReDim outputBytes(startTokenIndex To maxTokenIndex)
        'outputIndex is incremented by 2 whenever a value is added so initialize at -2.
        outputIndex = (-1 * BYTE_OFFSET) + startTokenIndex
        
        For iMainLoop = startTokenIndex To maxTokenIndex Step BYTE_OFFSET
        
            Select Case json(iMainLoop)
                
                Case backslash
                    ' \ encountered.
                    groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        iMainLoop = iMainLoop + BYTE_OFFSET
                        If iMainLoop > UBound(json) Then Exit Do
                    Loop Until json(iMainLoop) <> backslash
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    
                    If countOfBackslashToKeep > 0 Then
                        ' Write backslashes to output array.
                        For outputIndex = outputIndex + BYTE_OFFSET To outputIndex + (2 * countOfBackslashToKeep) Step BYTE_OFFSET
                            outputBytes(outputIndex) = backslash
                        Next outputIndex
                        ' For loops will exceed boundary on last loop so reverse by the value of step.
                        outputIndex = outputIndex - BYTE_OFFSET
                    End If
                    
                    ignoreSpecialCharacter = groupedBackslashCount Mod 2 = 0
                    
                    If iMainLoop <= UBound(json) Then
                    
                        outputIndex = outputIndex + BYTE_OFFSET
                        
                        If ignoreSpecialCharacter Then
                            outputBytes(outputIndex) = json(iMainLoop)
                        Else
                            Select Case json(iMainLoop)
                                Case 98 '\b
                                    outputBytes(outputIndex) = 8    'backspace
                                Case 110 '\n
                                    outputBytes(outputIndex) = 10   'newline
                                Case 116 '\t
                                    outputBytes(outputIndex) = 9    'tab
                                Case quote, 47, 0, 39 '\" or \/ or \0 or \'
                                    outputBytes(outputIndex) = json(iMainLoop)
                                Case 114 '\r
                                    outputBytes(outputIndex) = 13   'carriage return
                                Case 102 '\f
                                    outputBytes(outputIndex) = 12   ' Form feed
                                Case 117 '\u
                                    'Is unicode get 4 bytes after u and calculate.
                                    ' append &h
                                    For tempCount = 1 To 4
                                        iMainLoop = iMainLoop + BYTE_OFFSET
                                        uniCode((tempCount - 1) * 2) = json(iMainLoop)
                                    Next tempCount
                                    ' To Do properly parse utf characters.
                                    outputBytes(outputIndex) = Val("&h" & CStr(uniCode))

                                Case Else
                                    On Error GoTo 0
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                        
                    ElseIf Not ignoreSpecialCharacter Then
                        On Error GoTo 0
                        Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Unescaped backslash at end of string: '\'"
                    End If
                'Case 47
                    'Err.Raise INVALID_JSON_FORMAT_ERROR, "Pro bvcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped solidius: '/'"
                Case 10, 13, 9, 8, 12
                    'line feed,carriage return,tab,backspace,form feed
                Case Else
                    outputIndex = outputIndex + BYTE_OFFSET
                    outputBytes(outputIndex) = json(iMainLoop)
                    If json(iMainLoop) = quote Then unescapedQuoteCount = unescapedQuoteCount + 1
            End Select
            
        Next iMainLoop

        ' If there are 2 unescaped quotes in the byte array and they occur at the end, then byte segement is a string.
        If unescapedQuoteCount = 2 And (outputBytes(LBound(outputBytes)) = quote And outputBytes(outputIndex) = quote) Then
            'Value is a string.
            outputBytes = GetArraySubset(outputBytes, LBound(outputBytes) + BYTE_OFFSET, outputIndex - 1)
            isValueContainedWithinUnescapedQuotes = True
            
        ElseIf unescapedQuoteCount >= 1 Then
            ' Ensure quality of string value. ie don't allow "apple to be excepted as a string.
            On Error GoTo 0
            Err.Raise INVALID_JSON_FORMAT_ERROR, "Invalid Format: Unescaped Quote present in string: " & vbNewLine & CStr(GetArraySubset(outputBytes, LBound(outputBytes), outputIndex + 1))
        Else
            ' + 1 to encapsulate character byte.
            ReDim Preserve outputBytes(LBound(outputBytes) To outputIndex + BYTE_OFFSET - 1)
        End If
    Else
        outputBytes = GetArraySubset(json, startTokenIndex, maxTokenIndex)
    End If

    ProcessSpecialCharactersInJsonStringValue = outputBytes
    Exit Function

PropagateError:
    PropagateError Err, "ProcessSpecialCharactersInJsonStringValue"
    
End Function
Private Function TryFindFirstAndLastNonWhitespace(json() As Byte, startIndex&, maxIndex&, leftByteStart&, rightNonNullByte&) As Boolean
'========================================================================================================================
'   Summary:
'       Searches for non [ space,newlines,line feeds and carriage returns ] from both ends of [json()] if they exist.
'       leftByteStart and rightNonNullByte will be set to the relevant non-white space bytes.
'============================================================================================================
    Dim iCountRight&, iCountLeft&, _
    cleaningLeft As Boolean, cleaningRight As Boolean, iCount As Byte, errorLeft As Boolean, _
    errorRight As Boolean, cleanOnLeft As Boolean, cleanOnRight As Boolean

    iCountLeft = startIndex
    'Max index should be the non null portion of a string byte.
    iCountRight = maxIndex

    cleaningLeft = True: cleaningRight = True

    Do While cleaningLeft Or cleaningRight

        If cleaningLeft Then
            On Error GoTo CatchLeft
            
            Select Case json(iCountLeft)
                Case 32, 9, 10, 13
                    iCountLeft = iCountLeft + BYTE_OFFSET
                    cleaningLeft = Not iCountLeft > iCountRight And iCountLeft <= maxIndex
                    cleanOnLeft = True
                Case Else
                    cleaningLeft = False
                    leftByteStart = iCountLeft
            End Select
            
        End If
Try_Right:
        If cleaningRight Then
            On Error GoTo CatchRight
            Select Case json(iCountRight)
                Case 32, 9, 10, 13 '12
                    iCountRight = iCountRight - BYTE_OFFSET
                    cleaningRight = Not iCountRight < iCountLeft And iCountRight >= startIndex
                    cleanOnRight = True
                Case Else
                    cleaningRight = False
                    rightNonNullByte = iCountRight
            End Select
        End If
NextLoop:
    Loop

    If (errorLeft And errorRight) Or (iCountLeft > iCountRight) Then
        'Byte sequence is just whitespace characters
        TryFindFirstAndLastNonWhitespace = False
    Else
        TryFindFirstAndLastNonWhitespace = True
    End If

    Exit Function
'
CatchLeft:
    cleaningLeft = False
    errorLeft = True
    Resume Try_Right
CatchRight:
    cleaningRight = False
    errorRight = True
    Resume NextLoop
End Function
Private Function LocationOfCharacterNotBetweenTextTokens(ByRef query() As Byte, character$, startIndex&, maxIndex&) As Long
'===============================================================================================================
' Summary
'   Search for the next [character] not contained within quotes.
' Remarks:
'   [character] within escaped quotes will be ignored.
'===============================================================================================================
    Dim characterLocationBaseOne&, leftQuoteIndexBaseOne&, _
    rightQuoteIndexBaseOne&, adjustLeftQuote As Long, isNotEscapedQuote As Boolean, locationOfQuoteBaseOne&, updateCharacterlocation As Boolean
    
    Const quote$ = """", backSlashCode As Byte = 92
    On Error GoTo Return_Default
    
    Dim lboundMinusOne&: lboundMinusOne = LBound(query) - 1
    ' Assign to index base 1.
    rightQuoteIndexBaseOne = startIndex - lboundMinusOne - BYTE_OFFSET
    updateCharacterlocation = True
    Do
        If updateCharacterlocation Then characterLocationBaseOne = InStrB(rightQuoteIndexBaseOne + BYTE_OFFSET, query, character)

        leftQuoteIndexBaseOne = rightQuoteIndexBaseOne

        For adjustLeftQuote = -1 To 0

            locationOfQuoteBaseOne = leftQuoteIndexBaseOne
            Do
                locationOfQuoteBaseOne = InStrB(locationOfQuoteBaseOne + BYTE_OFFSET, query, quote)
                If Not locationOfQuoteBaseOne = 0 Then
                    ' Set to TRUE if previous character in the string isn't a backslash.
                    isNotEscapedQuote = query(locationOfQuoteBaseOne - BYTE_OFFSET + lboundMinusOne) <> backSlashCode
                End If
            Loop Until isNotEscapedQuote Or locationOfQuoteBaseOne = 0
            
            If adjustLeftQuote Then
                leftQuoteIndexBaseOne = locationOfQuoteBaseOne
            Else
                rightQuoteIndexBaseOne = locationOfQuoteBaseOne
            End If
            
        Next adjustLeftQuote
        ' Update if sandwhiched between quotes.
        updateCharacterlocation = (characterLocationBaseOne > leftQuoteIndexBaseOne And characterLocationBaseOne < rightQuoteIndexBaseOne)

    Loop Until (characterLocationBaseOne < leftQuoteIndexBaseOne And Not updateCharacterlocation) Or characterLocationBaseOne = 0 Or characterLocationBaseOne > maxIndex Or leftQuoteIndexBaseOne = 0
    
    Select Case characterLocationBaseOne
        Case Is > maxIndex, 0
            LocationOfCharacterNotBetweenTextTokens = 0
        Case Else
            If Not updateCharacterlocation Then
                LocationOfCharacterNotBetweenTextTokens = characterLocationBaseOne + lboundMinusOne
            Else
                LocationOfCharacterNotBetweenTextTokens = 0
            End If
    End Select
    
    Exit Function
Return_Default:
    LocationOfCharacterNotBetweenTextTokens = 0
End Function

Private Sub PropagateError(e As ErrObject, procedureName$, Optional moreDetails$ = vbNullString)
    
    Dim firstPropagation As Boolean, sourceParts$()
    
    Const delim$ = ": "
    
    With e
        AddParentToErrSource e, procedureName
        If LenB(moreDetails) > 0 Then .description = moreDetails & vbNewLine & .description
        .Raise .number, .source, .description
    End With
    
End Sub

Private Function GetArraySubset(query() As Byte, startIndex&, maxIndex&) As Byte()

    Dim outputA() As Byte, iCount&
    ' Check to make sure section to copy has enough bytes. 2 bytes per character. Ex: 3char string  = 6 bytes
    If (maxIndex - startIndex) Mod 2 = 0 And BYTE_OFFSET = 2 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, "GetArraySubset", "Attempted to create a string from an uneven length byte array."
    End If
    
    ReDim outputA(startIndex To maxIndex)
    
    For iCount = startIndex To maxIndex Step BYTE_OFFSET
        outputA(iCount) = query(iCount)
    Next iCount
    GetArraySubset = outputA
    
End Function
Private Sub AddParentToErrSource(e As ErrObject, parentName$)
    
    Dim sourceParts$()
    Const delim$ = ": "
    
    With e
    
        If InStrB(1, .source, delim) = 0 Then
            .source = "[" & .source & "]" & delim & parentName
        Else
            sourceParts = Split(.source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .source = Join(sourceParts, delim)
        End If
                    
    End With
    
End Sub
Private Sub PreProcessData(ByRef jsonBytes() As Byte)
    
    Dim i&, nullCount&, unicodeBytes As Boolean, offset&, currentUBound&, simpleMap As Boolean
    
    unicodeBytes = True
    ' Unicode byte arrays are even length. If odd then not unicode.
    If (UBound(jsonBytes) - LBound(jsonBytes) + 1) Mod 2 = 1 Then unicodeBytes = False

    If unicodeBytes Then
        ' Ensure all bytes are in range.
        simpleMap = True
        For i = LBound(jsonBytes) To UBound(jsonBytes)
            If jsonBytes(i) = 0 Then nullCount = nullCount + 1
            If jsonBytes(i) > 255 Then
                simpleMap = False
                unicodeBytes = False
                Exit For
            End If
        Next i
        If nullCount <> (UBound(jsonBytes) - LBound(jsonBytes) + 1) / 2 Then unicodeBytes = False
    End If
    
    If Not unicodeBytes Then
        
        If simpleMap Then
            ReDim Preserve jsonBytes(LBound(jsonBytes) To -LBound(jsonBytes) + (UBound(jsonBytes) - LBound(jsonBytes) + 1) * 2)
            
            offset = (UBound(jsonBytes) / 2) - 1

            For i = UBound(jsonBytes) - 2 To LBound(jsonBytes) + 1 Step -2
                jsonBytes(i) = jsonBytes(offset)
                offset = offset - 1
            Next i
        Else
            'jsonbytes=DecodeUTF8(jsonbytes,true)
        End If

    End If

    'BYTE_OFFSET = 2 'IIf(unicodeBytes, 2, 1)
    
End Sub

'Public Function DecodeUTF8(ByRef utf8Str As String, _
'                  Optional ByVal raiseErrors As Boolean = False) As String
'
''Function transcoding an UTF-8 encoded string to the VBA-native UTF-16-LE
''Author: Guido Witt-Dörring, https://stackoverflow.com/a/75787820/12287457
''                            https://github.com/guwidoe/VBA-StringTools
'
'    Const methodName As String = "DecodeUTF8"
'    Dim i As Long
'    Dim numBytesOfCodePoint As Byte
'
'    Static numBytesOfCodePoints(0 To 255) As Byte
'    Static mask(2 To 4) As Long
'    Static minCp(2 To 4) As Long
'
'    If numBytesOfCodePoints(0) = 0 Then
'        For i = &H0& To &H7F&: numBytesOfCodePoints(i) = 1: Next i '0xxxxxxx
'        '110xxxxx - C0 and C1 are invalid (overlong encoding)
'        For i = &HC2& To &HDF&: numBytesOfCodePoints(i) = 2: Next i
'        For i = &HE0& To &HEF&: numBytesOfCodePoints(i) = 3: Next i '1110xxxx
'       '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
'        For i = &HF0& To &HF4&: numBytesOfCodePoints(i) = 4: Next i
'        For i = 2 To 4: mask(i) = (2 ^ (7 - i) - 1): Next i
'        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
'    End If
'
'    Dim codepoint As Long
'    Dim currByte As Byte
'    Dim utf8() As Byte:  utf8 = utf8Str
'    Dim utf16() As Byte: ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
'    Dim j As Long:       j = 0
'    Dim k As Long
'
'    i = LBound(utf8)
'    Do While i <= UBound(utf8)
'        codepoint = utf8(i)
'        numBytesOfCodePoint = numBytesOfCodePoints(codepoint)
'
'        If numBytesOfCodePoint = 0 Then
'            If raiseErrors Then Err.Raise 5, methodName, "Invalid byte"
'            GoTo insertErrChar
'        ElseIf numBytesOfCodePoint = 1 Then
'            utf16(j) = codepoint
'            j = j + 2
'        ElseIf i + numBytesOfCodePoint - 1 > UBound(utf8) Then
'            If raiseErrors Then Err.Raise 5, methodName, _
'                    "Incomplete UTF-8 codepoint at end of string."
'            GoTo insertErrChar
'        Else
'            codepoint = utf8(i) And mask(numBytesOfCodePoint)
'
'            For k = 1 To numBytesOfCodePoint - 1
'                currByte = utf8(i + k)
'
'                If (currByte And &HC0&) = &H80& Then
'                    codepoint = (codepoint * &H40&) + (currByte And &H3F)
'                Else
'                    If raiseErrors Then _
'                        Err.Raise 5, methodName, "Invalid continuation byte"
'                    GoTo insertErrChar
'                End If
'            Next k
'            'Convert the Unicode codepoint to UTF-16LE bytes
'            If codepoint < minCp(numBytesOfCodePoint) Then
'                If raiseErrors Then Err.Raise 5, methodName, "Overlong encoding"
'                GoTo insertErrChar
'            ElseIf codepoint < &HD800& Then
'                utf16(j) = codepoint And &HFF&
'                utf16(j + 1) = codepoint \ &H100&
'                j = j + 2
'            ElseIf codepoint < &HE000& Then
'                If raiseErrors Then Err.Raise 5, methodName, _
'                "Invalid Unicode codepoint.(Range reserved for surrogate pairs)"
'                GoTo insertErrChar
'            ElseIf codepoint < &H10000 Then
'                If codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
'                utf16(j) = codepoint And &HFF&
'                utf16(j + 1) = codepoint \ &H100&
'                j = j + 2
'            ElseIf codepoint < &H110000 Then 'Calculate surrogate pair
'                Dim m As Long:           m = codepoint - &H10000
'                Dim loSurrogate As Long: loSurrogate = &HDC00& Or (m And &H3FF)
'                Dim hiSurrogate As Long: hiSurrogate = &HD800& Or (m \ &H400&)
'
'                utf16(j) = hiSurrogate And &HFF&
'                utf16(j + 1) = hiSurrogate \ &H100&
'                utf16(j + 2) = loSurrogate And &HFF&
'                utf16(j + 3) = loSurrogate \ &H100&
'                j = j + 4
'            Else
'                If raiseErrors Then Err.Raise 5, methodName, _
'                        "Codepoint outside of valid Unicode range"
'insertErrChar:  utf16(j) = &HFD
'                utf16(j + 1) = &HFF
'                j = j + 2
'
'                If numBytesOfCodePoint = 0 Then numBytesOfCodePoint = 1
'            End If
'        End If
'nextCp: i = i + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
'    Loop
'    DecodeUTF8 = MidB$(utf16, 1, j)
'End Function

