' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private BYTE_OFFSET&

Option Explicit
Public Function ParseJsonString(ByRef json$) As Object
    
    Dim jsonBytes() As Byte
    
    jsonBytes = json
    DetermineByteOffset jsonBytes
    Set ParseJsonString = ParseJSON(jsonBytes)
    
End Function
Public Function ParseJsonBytes(ByRef jsonBytes() As Byte) As Object
    
    DetermineByteOffset jsonBytes
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

    Dim validInput As Boolean, keyForValue$, jsonContainer As Object, keyValueDelimiterTokenIndex&, nextTokenDelimiterIndex&, keyStart&, keyEnd&, valueStart&, valueEnd&, _
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
                            ' Ensure that all quotes except enclosing quotes are escaped.
                            If Not IsUnescapedQuoteInText(json, keyStart, keyEnd) Then
                                
                                ' Ensure Key ends in an unescaped quote and doesn't overlap with value.
                                If ((json(keyEnd) = quote And json(keyEnd - BYTE_OFFSET) <> backslash And json(keyEnd) = 34) And (keyEnd < valueStart)) Then
                                    
                                    keyEnd = keyEnd + (BYTE_OFFSET - 1)
                                    keyForValue = ProcessSpecialCharactersInJsonStringValue(json, keyStart, keyEnd)
                                    keyAvailable = True

                                Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Key-Value pair is invalid."
                                End If
                            Else
                                 Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Unescaped quote detected in key: " & CStr(GetArraySubset(json, keyStart, keyEnd + (BYTE_OFFSET - 1)))
                            End If
                        Else
                            Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Invalid Format: Key-Value delimiter not found."
                        End If
                    Else
                        Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Invalid Format: Key doesn't start with a quote: " & CStr(GetArraySubset(json, keyStart, valueEnd + (BYTE_OFFSET - 1)))
                    End If
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Invalid Format: Expected Key-Value pair not found."
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
                    'Adjust by 2 to remove token delimiters.
                    valueAvailable = TryFindFirstAndLastNonWhitespace(json, currentTokenIndex + BYTE_OFFSET, nextTokenDelimiterIndex - BYTE_OFFSET, valueStart, valueEnd)
                End If
                ' Ends are currently set to the first byte in a 2 byte character. Set them to the second to encapsulate.
                valueEnd = valueEnd + (BYTE_OFFSET - 1)

                If valueAvailable And ((json(valueStart) = leftCurly And json(valueEnd - (BYTE_OFFSET - 1)) = rightCurly) Or (json(valueStart) = leftSquare And json(valueEnd - (BYTE_OFFSET - 1)) = rightSquare)) Then
                    If parsingJsonObject Then
                        jsonContainer.Add IIf(keyAvailable, keyForValue, unIdentifiedCount), ParseJSON(json, valueStart, valueEnd, recursionDepth, keyForValue)
                    Else
                        jsonContainer.Add ParseJSON(json, valueStart, valueEnd, recursionDepth, IIf(keyAvailable, keyForValue, unIdentifiedCount))
                    End If
                Else
                    ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                    ' Throws an error if any other unescaped quotes.
                    If valueAvailable Then
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
                                Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJSON", "Invalid format: Value isn't numeric or enclosed in an unescaped string: " & value
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
    
    Dim nextCurlyClose&, nextCurlyOpen&, previousClose&, tokenIndicator$(1), findEndOfArrayOrObject As Boolean, nextTokenDelimiterIndex&
    
    Dim lboundMinusOne&: lboundMinusOne = LBound(json) - 1
    
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
            findEndOfArrayOrObject = True
            
        ElseIf nextSquareOpen < nextValidCommaTokenIndexBaseOne And nextSquareOpen > 0 And (nextSquareOpen < nextCurlyOpen Or nextCurlyOpen = 0) Then
            
            tokenIndicator(0) = leftSquare: tokenIndicator(1) = rightSquare
            nextOpen = nextSquareOpen: nextClose = nextSquareClose
            findEndOfArrayOrObject = True
            
        End If
    
        If findEndOfArrayOrObject Then
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
            Err.Raise INVALID_JSON_FORMAT_ERROR, "GetNextTokenDelimiterIndex", "Invalid JSON format. Calculated token index doesn't reference a value delimiter."
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
    
    specialChars(0) = Chr(34): specialChars(1) = "\"
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
                
                Case 92
                    ' \ encountered.
                    groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        iMainLoop = iMainLoop + BYTE_OFFSET
                        If iMainLoop > UBound(json) Then Exit Do
                    Loop Until json(iMainLoop) <> 92
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    
                    If countOfBackslashToKeep > 0 Then
                        ' Write backslashes to output array.
                        For outputIndex = outputIndex + BYTE_OFFSET To outputIndex + (2 * countOfBackslashToKeep) Step BYTE_OFFSET
                            outputBytes(outputIndex) = 92
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
                                    'backspace
                                    outputBytes(outputIndex) = 8
                                Case 110 '\n
                                    'newline
                                    outputBytes(outputIndex) = 10
                                Case 116 '\t
                                    'tab
                                    outputBytes(outputIndex) = 9
                                Case 34, 47, 0, 39 '\" or \/ or \0 or \'
                                    outputBytes(outputIndex) = json(iMainLoop)
                                Case 114 '\r
                                    'carriage return
                                    outputBytes(outputIndex) = 13
                                Case 102 '\f
                                    ' Form feed
                                    outputBytes(outputIndex) = 12
                                Case 117 '\u
                                    'Is unicode get 4 bytes after u and calculate.
                                    ' append &h
                                    For tempCount = 1 To 4
                                        iMainLoop = iMainLoop + BYTE_OFFSET
                                        uniCode((tempCount - 1) * 2) = json(iMainLoop)
                                    Next tempCount
                                    outputBytes(outputIndex) = Val("&h" & CStr(uniCode))
                                Case Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ProcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                        
                    ElseIf Not ignoreSpecialCharacter Then
                        Err.Raise INVALID_JSON_FORMAT_ERROR, "ProcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped backslash at end of string: '\'"
                    End If
                'Case 47
                    'Err.Raise INVALID_JSON_FORMAT_ERROR, "Pro bvcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped solidius: '/'"
                Case 10, 13, 9, 8, 12
                    'line feed,carriage return,tab,backspace,form feed
                Case Else
                    outputIndex = outputIndex + BYTE_OFFSET
                    outputBytes(outputIndex) = json(iMainLoop)
                    If json(iMainLoop) = 34 Then unescapedQuoteCount = unescapedQuoteCount + 1
            End Select
            
        Next iMainLoop

        'Ensure that the quotes are on the end.
        If unescapedQuoteCount = 2 And (outputBytes(LBound(outputBytes)) = 34 And outputBytes(outputIndex) = 34) Then
            'Value is a string.
            outputBytes = GetArraySubset(outputBytes, LBound(outputBytes) + BYTE_OFFSET, outputIndex - 1)
            isValueContainedWithinUnescapedQuotes = True

        ElseIf unescapedQuoteCount >= 1 Then
            Err.Raise INVALID_JSON_FORMAT_ERROR, "ProcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped Quote present in string: " & vbNewLine & CStr(GetArraySubset(outputBytes, LBound(outputBytes), outputIndex + 1))
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
Private Function TryFindFirstAndLastNonWhitespace(json() As Byte, startIndex&, maxIndex&, first&, last&) As Boolean
'========================================================================================================================
'   Summary:
'       Removes space,newlines,line feeds and carriage returns from both ends of [query] if they exist.
'   first and last will be set to the relevant non-white space bytes.
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
                    first = iCountLeft
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
                    last = iCountRight
            End Select
        End If
NextLoop:
    Loop
'   iCountLeft will always be 2 more than what you want.
    'first = IIf(cleanOnLeft, iCountLeft - byte_offset, startIndex)
    'last = IIf(cleanOnRight, iCountRight - byte_offset, maxIndex)

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
Private Function IsUnescapedQuoteInText(json() As Byte, ByVal startIndexQuote&, ByVal maxIndexQuote&) As Boolean
'   Summary:
'       Determine if there is an unescaped quote in [json] or if the last quote has been escaped.
'   Notes:
'       StartIndex and Max index should be locations of enclosing quotes ie: value =34.
    On Error GoTo Propagate
    Dim unescapedQuoteFound As Boolean, lboundMinusOne&: lboundMinusOne = LBound(json) - 1
    Const quote$ = """", slash$ = "\"
    
    startIndexQuote = startIndexQuote - lboundMinusOne
    maxIndexQuote = maxIndexQuote - lboundMinusOne

    Do
        startIndexQuote = InStrB(startIndexQuote + BYTE_OFFSET, json, quote)
        If startIndexQuote <> 0 Then
            Select Case startIndexQuote
                Case maxIndexQuote
                    'Quote at end of string. True if escaped.
                    unescapedQuoteFound = json(startIndexQuote - BYTE_OFFSET) = 92
                Case Else
                    unescapedQuoteFound = json(startIndexQuote - BYTE_OFFSET) <> 92
            End Select
        End If
    Loop While startIndexQuote <> 0 And startIndexQuote <= maxIndexQuote And Not unescapedQuoteFound
    Exit Function
Propagate:
    PropagateError Err, "IsUnescapedQuoteInText"
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
                ' Set to true if previous character in the string isn't a backslash.
                If Not locationOfQuoteBaseOne = 0 Then
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
Private Sub DetermineByteOffset(ByRef jsonBytes() As Byte)
    
    Dim I&, nullCount&, unicodeBytes As Boolean
    
    For I = LBound(jsonBytes) + 1 To UBound(jsonBytes) Step 2
    
        If jsonBytes(I) = 0 And I <> UBound(jsonBytes) Then
            unicodeBytes = True
            Exit For
        ElseIf jsonBytes(I) = 0 Then
            nullCount = nullCount + 1
            If I = UBound(jsonBytes) Or nullCount >= 1 Then
                unicodeBytes = True
                Exit For
            End If
        End If
    Next I
    
    BYTE_OFFSET = IIf(unicodeBytes, 2, 1)
    
End Sub
