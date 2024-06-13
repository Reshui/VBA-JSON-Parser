' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Option Explicit
Public Function ParseJSON(ByRef json$) As Object
    
    Dim jsonBytes() As Byte
    
    jsonBytes = json
    Set ParseJSON = ParseJsonBytes(jsonBytes)
    
End Function
Public Function ParseJsonBytes(ByRef json() As Byte, Optional ByRef currentTokenIndex As Long = 1, Optional maxIndex& = -1, Optional ByVal recursionDepth& = -1, Optional ByVal arrayOrObjectName$ = vbNullString, Optional ByRef utf8Bytes As Boolean = False) As Object
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
    Dim keyAvailable As Boolean, value$, unIdentifiedCount&, parsingJsonArray As Boolean, parsingJsonObject As Boolean, finalValueBytes() As Byte

    Dim validInput As Boolean, keyForValue$, currentDCTN As Object, keyValueDelimiterTokenIndex&, nextTokenDelimiterIndex&, keyStart&, keyEnd&, valueStart&, valueEnd&
    
    Const quote As Byte = 34, backslash As Byte = 92, rightCurly As Byte = 125, leftCurly As Byte = 123, colon$ = ":", leftSquare As Byte = 91, rightSquare As Byte = 93 ', arrayPattern$ = "[[]*[]]", objectPattern$ = "{*}", quoteEnclosedPattern$ = """*"""

    recursionDepth = recursionDepth + 1: unIdentifiedCount = -1
'
    If recursionDepth = 0 Then
        maxIndex = UBound(json)
        currentTokenIndex = LBound(json)
    End If

    On Error GoTo CATCH_FAILURE
    
    parsingJsonObject = (json(currentTokenIndex) = leftCurly And json(maxIndex - 1) = rightCurly)
    parsingJsonArray = (json(currentTokenIndex) = leftSquare And json(maxIndex - 1) = rightSquare)
    
    If parsingJsonObject Or parsingJsonArray Then
        'Max index - 1 should be the last non-null character in the range.
        Do While nextTokenDelimiterIndex < maxIndex - 1

            nextTokenDelimiterIndex = GetNextTokenDelimiterIndex(json, currentTokenIndex, maxIndex - 1)

            If parsingJsonObject Then
                'Ex: {"apple":45}
                ' If non whitespace characters exist then record their location; otherwise throw an error.
                If TryFindFirstAndLastNonWhitespace(json, currentTokenIndex + 2, nextTokenDelimiterIndex - 2, keyStart, valueEnd) Then
                    ' Does key start with a quote?
                    If json(keyStart) = quote Then
                        
                        keyValueDelimiterTokenIndex = LocationOfCharacterNotBetweenTextTokens(json, colon, keyStart, valueEnd)
                        ' Is there a colon outside of quotes?
                        If keyValueDelimiterTokenIndex > 0 Then
                            ' Find boundaries of key using location of key-value delimiter.
                            Call TryFindFirstAndLastNonWhitespace(json, keyStart, keyValueDelimiterTokenIndex - 2, keyStart, keyEnd)
                            ' Find boundaries of value.
                            Call TryFindFirstAndLastNonWhitespace(json, keyValueDelimiterTokenIndex + 2, nextTokenDelimiterIndex - 2, valueStart, valueEnd)
                            ' Ensure that all quotes except enclosing quotes are escaped.
                            If Not IsUnescapedQuoteInText(json, keyStart, keyEnd) Then
                                ' Ensure Key ends in an unescaped quote and doesn't overlap with value.
                                If ((json(keyEnd) = quote And json(keyEnd - 2) <> backslash And json(keyEnd) = 34) And (keyEnd < valueStart)) Then
                                    keyAvailable = True
                                Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Key-Value pair is invalid."
                                End If
                            Else
                                 Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Unescaped quote detected in key: " & CStr(GetArraySubset(json, keyStart, keyEnd + 1))
                            End If
                        Else
                            Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Invalid Format: Key-Value delimiter not found."
                        End If
                    Else
                        Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Invalid Format: Key doesn't start with a quote: " & CStr(GetArraySubset(json, keyStart, valueEnd + 1))
                    End If
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Invalid Format: Expected Key-Value pair not found."
                End If
            End If

            If parsingJsonArray Or (keyAvailable And parsingJsonObject) Then

                If currentDCTN Is Nothing Then
                    If parsingJsonObject Then
                        #If Mac Then
                            Set currentDCTN = New Dictionary
                        #Else
                            Set currentDCTN = CreateObject("Scripting.Dictionary")
                        #End If
                    Else
                        Set currentDCTN = New Collection
                    End If
                End If
                            
                If parsingJsonArray Then
                    unIdentifiedCount = unIdentifiedCount + 1
                    'Adjust by 2 to remove token delimiters.
                    Call TryFindFirstAndLastNonWhitespace(json, currentTokenIndex + 2, nextTokenDelimiterIndex - 2, valueStart, valueEnd)
                Else
                    ' Ends are currently set to the first byte in a 2 byte character. Set them to the second to encapsulate.
                    keyEnd = keyEnd + 1
                    ' Adjust by 2 to remove quotes.
                    keyForValue = ProcessSpecialCharactersInJsonStringValue(json, keyStart + 2, keyEnd - 2)
                End If
                ' Ends are currently set to the first byte in a 2 byte character. Set them to the second to encapsulate.
                valueEnd = valueEnd + 1

                If (json(valueStart) = leftCurly And json(valueEnd - 1) = rightCurly) Or (json(valueStart) = leftSquare And json(valueEnd - 1) = rightSquare) Then
                    If parsingJsonObject Then
                        currentDCTN.Add IIf(keyAvailable, keyForValue, unIdentifiedCount), ParseJsonBytes(json, valueStart, valueEnd, recursionDepth, keyForValue)
                    Else
                        currentDCTN.Add ParseJsonBytes(json, valueStart, valueEnd, recursionDepth, keyForValue)
                    End If
                Else
                    
                    value = GetArraySubset(json, valueStart, valueEnd) 'ProcessSpecialCharactersInJsonStringValue(json, valueStart, valueEnd)
                    
                    If IsNumeric(value) Then
                        If parsingJsonObject Then
                            currentDCTN.Add keyForValue, IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
                        Else
                            currentDCTN.Add IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
                        End If
                    Else
                        Select Case value
                            Case "true", "false"
                                If parsingJsonObject Then
                                    currentDCTN.Add keyForValue, CBool(value)
                                Else
                                    currentDCTN.Add CBool(value)
                                End If
                            Case "null", vbNullString
                                If parsingJsonObject Then
                                    currentDCTN.Add keyForValue, Null
                                Else
                                    currentDCTN.Add Null
                                End If
                            Case Else
                                ' Ensure that the value is enclosed in quotes to denote a valid string.
                                If (json(valueStart) = quote And json(valueEnd - 1) = quote) Then
                                    If Not IsUnescapedQuoteInText(json, valueStart, valueEnd - 1) Then
                                        ' Remove the enclosing quotes
                                        value = ProcessSpecialCharactersInJsonStringValue(json, valueStart + 2, valueEnd - 2)
                                        
                                        If parsingJsonObject Then
                                            currentDCTN.Add keyForValue, value
                                        Else
                                            currentDCTN.Add value
                                        End If
                                    Else
                                        Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Invalid format: Unescaped quote within value body or escaped last quote. : " & value
                                    End If
                                Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, "ParseJsonBytes", "Invalid format: Value isn't numeric and should be enclosed in unescaped quotes. : " & value
                                End If
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
    Set ParseJsonBytes = currentDCTN
    Exit Function
CATCH_FAILURE:
    'Stop: Resume
    PropagateError Err, "ParseJsonBytes", "Name: '" & arrayOrObjectName & "' Depth: " & recursionDepth & " Item: " & IIf(keyAvailable, keyForValue, unIdentifiedCount)
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
    
    nextValidCommaTokenIndexBaseOne = LocationOfCharacterNotBetweenTextTokens(json, comma, currentTokenIndex + 2, maxTokenIndex)

    If nextValidCommaTokenIndexBaseOne > 0 Then nextValidCommaTokenIndexBaseOne = nextValidCommaTokenIndexBaseOne - lboundMinusOne
    '===========================================================================================================================
    ' Determine which of [ or { is greater than currentTokenIndex AND appears first.
    ' If nextValidCommaTokenIndexBaseOne is greater than the first array or object token start, then find end token for the array or object + 1.
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
            nextClose = InStrB(previousClose + 2, json, tokenIndicator(1))
            nextOpen = InStrB(nextOpen + 2, json, tokenIndicator(0))
        Loop Until nextOpen > previousClose Or nextOpen = 0
        ' Set tempIndexBaseOne to the next character after the end of the array or object that triggered loop.
        tempIndexBaseOne = previousClose + 2
        nextValidCommaTokenIndexBaseOne = LocationOfCharacterNotBetweenTextTokens(json, comma, tempIndexBaseOne + lboundMinusOne, maxTokenIndex)
        If nextValidCommaTokenIndexBaseOne > 0 Then nextValidCommaTokenIndexBaseOne = nextValidCommaTokenIndexBaseOne - lboundMinusOne
    End If
    '===========================================================================================================================
    nextTokenDelimiterIndex = nextValidCommaTokenIndexBaseOne + IIf(nextValidCommaTokenIndexBaseOne > 0, lboundMinusOne, 0)

    Select Case nextTokenDelimiterIndex
        Case 0, currentTokenIndex, Is > maxTokenIndex
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
Private Function ProcessSpecialCharactersInJsonStringValue(ByRef json() As Byte, startTokenIndex&, maxTokenIndex&) As Byte()
'========================================================================================================================================
'Summary: Converts special characters that are lead by a \.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes backslash
'========================================================================================================================================
    Dim outputBytes() As Byte, iMainLoop&, outputIndex&, ignoreSpecialCharacter As Boolean, lboundMinusOne&, _
    groupedBackslashCount&, backslashesToKeepCount&, tempCount&, uniCode(8) As Byte, processBytes As Boolean, specialChars(6) As String
    
    specialChars(0) = "\": specialChars(1) = """"
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
        outputIndex = -2 + startTokenIndex
        
        For iMainLoop = startTokenIndex To maxTokenIndex Step 2
        
            Select Case json(iMainLoop)
                
                Case 92
                    ' \ encountered.
                    groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        iMainLoop = iMainLoop + 2
                        If iMainLoop > UBound(json) Then Exit Do
                    Loop Until json(iMainLoop) <> 92
                    
                    ' Determine how many \ to keep.
                    backslashesToKeepCount = groupedBackslashCount \ 2
                    
                    If backslashesToKeepCount > 0 Then
                        ' Write backslashes to output array.
                        For outputIndex = outputIndex + 2 To outputIndex + (2 * backslashesToKeepCount) Step 2
                            outputBytes(outputIndex) = 92
                        Next outputIndex
                        outputIndex = outputIndex - 2
                    End If
                    
                    ignoreSpecialCharacter = groupedBackslashCount Mod 2 = 0
                    
                    If iMainLoop <= UBound(json) Then
                    
                        outputIndex = outputIndex + 2
                        
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
                                    'quote
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
                                        iMainLoop = iMainLoop + 2
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
                    'Err.Raise INVALID_JSON_FORMAT_ERROR, "ProcessSpecialCharactersInJsonStringValue", "Invalid Format: Unescaped solidius: '/'"
                Case 10, 13, 9, 8, 12
                    'line feed,carriage return,tab,backspace,form feed
                Case Else
                    outputIndex = outputIndex + 2
                    outputBytes(outputIndex) = json(iMainLoop)
            End Select
            
        Next iMainLoop
        
        ReDim Preserve outputBytes(LBound(outputBytes) To outputIndex + 1)
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
                    iCountLeft = iCountLeft + 2
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
                    iCountRight = iCountRight - 2
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
    'first = IIf(cleanOnLeft, iCountLeft - 2, startIndex)
    'last = IIf(cleanOnRight, iCountRight - 2, maxIndex)

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
        startIndexQuote = InStrB(startIndexQuote + 2, json, quote)
        If startIndexQuote <> 0 Then
            Select Case startIndexQuote
                Case maxIndexQuote
                    'Quote at end of string. True if escaped.
                    unescapedQuoteFound = json(startIndexQuote - 2) = 92
                Case Else
                    unescapedQuoteFound = json(startIndexQuote - 2) <> 92
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
    rightQuoteIndexBaseOne = startIndex - lboundMinusOne - 2
    updateCharacterlocation = True
    Do
        If updateCharacterlocation Then characterLocationBaseOne = InStrB(rightQuoteIndexBaseOne + 2, query, character)

        leftQuoteIndexBaseOne = rightQuoteIndexBaseOne

        For adjustLeftQuote = -1 To 0

            locationOfQuoteBaseOne = leftQuoteIndexBaseOne
            Do
                locationOfQuoteBaseOne = InStrB(locationOfQuoteBaseOne + 2, query, quote)
                ' Set to true if previous character in the string isn't a backslash.
                If Not locationOfQuoteBaseOne = 0 Then
                    isNotEscapedQuote = query(locationOfQuoteBaseOne - 2 + lboundMinusOne) <> backSlashCode
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
    If (maxIndex - startIndex) Mod 2 = 0 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, "GetArraySubset", "Attempted to create a string from an uneven length byte array."
    End If
    
    ReDim outputA(startIndex To maxIndex)
    
    For iCount = startIndex To maxIndex Step 2
        outputA(iCount) = query(iCount)
    Next iCount
    GetArraySubset = outputA
    
End Function

