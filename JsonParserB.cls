' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const BYTE_OFFSET As Byte = 2

Option Explicit
Public Function ParseJsonString(ByRef json$, canEraseString As Boolean) As Object
    
    Dim jsonBytes() As Byte
    
    jsonBytes = json
    If canEraseString Then json = vbNullString
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

Private Function ParseJSON(ByRef json() As Byte) As Object
''====================================================================================================================================================================================
'' Summary: Parse JSON byte data.
'' Inputs:
''   json: String to parse.
''   cursorByte: Index to start parsing from. Initialize at 1 or ensure it is a json token.
''   recursionDepth: Used for generating error messages.
''   arrayOrObjectName: Used for generating error messages.
'' Reference:
''   https://www.json.org/json-en.html
''====================================================================================================================================================================================
    Dim value$, unIdentifiedCount&, isValueJsonObject As Boolean, isValueJsonArray As Boolean, isParentJsonObject As Boolean, valueAvailable As Boolean

    Dim finalByteCursor&, keyForValue$, valueStartByte&, valueEndByte&, parentObject As Object, _
    childObject As Object, isValidatedString As Boolean, objectStack As New Collection, cursorByte&, nextDelimitingCursorByte&

    Const rightCurly As Byte = 125, leftCurly As Byte = 123, leftSquare As Byte = 91, rightSquare As Byte = 93
    
    Const cursorLimitKey$ = "MoshiM_UC@outlook.com"
    Const nextDelimitingParentCursorKey$ = "https://github.com/Reshui?tab=repositories"
    
    On Error GoTo CATCH_FAILURE

    if TryGetWhitespaceCleanedEnclosingBytes (json, LBound(json), UBound(json) - (BYTE_OFFSET - 1), cursorByte, finalByteCursor) then
    
        finalByteCursor = finalByteCursor - (BYTE_OFFSET - 1)        
        unIdentifiedCount = -1        
        
        If (json(cursorByte) = leftCurly And json(finalByteCursor) = rightCurly) Then
            isValueJsonObject = True
        ElseIf (json(cursorByte) = leftSquare And json(finalByteCursor) = rightSquare) Then
            isValueJsonArray = True
        End If

    end if

    If isValueJsonObject Or isValueJsonArray Then
        
        With objectStack
            If isValueJsonObject Then
                #If Mac Then
                    .Add New Dictionary
                #Else
                    .Add CreateObject("Scripting.Dictionary")
                #End If
            ElseIf isValueJsonArray Then
                .Add New Collection
            End If
            
            Set parentObject = .Item(1)
            Set childObject = .Item(1)
        End With
        
        AddValueToObject parentObject, addingNumber:=True, numberToAdd:=finalByteCursor, keyForValue:=cursorLimitKey
        AddValueToObject parentObject, addingNumber:=True, numberToAdd:=finalByteCursor, keyForValue:=nextDelimitingParentCursorKey
        
        Do While cursorByte < parentObject.Item(cursorLimitKey)

            Do While cursorByte < childObject.Item(cursorLimitKey)
                ' Finds the next token that encloses a value in conjunction with the current value of cursorByte.
                nextDelimitingCursorByte = GetNextTokenDelimiterIndex(json, cursorByte, childObject.Item(cursorLimitKey))
                ' nextDelimitingCursorByte = cursorByte will be true in situations like }}} because GetNextTokenDelimiterIndex increments cursorByte by 2 if it is a } or ]
                If Not nextDelimitingCursorByte = cursorByte Then
                
                    isParentJsonObject = (TypeName(childObject) <> "Collection")
                
                    If isParentJsonObject Then
                        GetKeyAndValueIndexes json, cursorByte, nextDelimitingCursorByte, valueStartByte, valueEndByte, valueAvailable, keyForValue
                    End If
        
                    If Not isParentJsonObject Or (LenB(keyForValue) > 0 And isParentJsonObject) Then
                                    
                        If Not isParentJsonObject Then
                            unIdentifiedCount = unIdentifiedCount + 1
                            'Adjust by BYTE_OFFSET to exclude token delimiters.
                            valueAvailable = TryGetWhitespaceCleanedEnclosingBytes(json, cursorByte + BYTE_OFFSET, nextDelimitingCursorByte - BYTE_OFFSET, valueStartByte, valueEndByte)
                        End If
        
                        'If enclosed in {} or [] then parse a new object or array.
                        isValueJsonObject = valueAvailable And (json(valueStartByte) = leftCurly And json(valueEndByte - (BYTE_OFFSET - 1)) = rightCurly)
                        isValueJsonArray = valueAvailable And (json(valueStartByte) = leftSquare And json(valueEndByte - (BYTE_OFFSET - 1)) = rightSquare)
    
                        If valueAvailable And (isValueJsonObject Or isValueJsonArray) Then
                            ' nextDelimitingCursorByte is boundary for the child object/array
                            ' one of either ] or ]
                            With objectStack
                                If isValueJsonObject Then
                                    #If Mac Then
                                        .Add New Dictionary
                                    #Else
                                        .Add CreateObject("Scripting.Dictionary")
                                    #End If
                                Else
                                    .Add New Collection
                                End If
                                
                                Set childObject = .Item(.count)
                                Set parentObject = .Item(.count - 1)
                            End With
                            ' Add key to limit byte range for object.
                            AddValueToObject childObject, addingNumber:=True, numberToAdd:=(valueEndByte - (BYTE_OFFSET - 1)), keyForValue:=cursorLimitKey
                            ' Byte to update cursorLocation to once array or object has been completed.
                            AddValueToObject childObject, addingNumber:=True, numberToAdd:=nextDelimitingCursorByte, keyForValue:=nextDelimitingParentCursorKey
                            ' Create heiarchy by nesting child within parent.
                            AddValueToObject parentObject, newChild:=childObject, keyForValue:=keyForValue
                            ' valueStartByte is one of either { or [
                            cursorByte = valueStartByte
                        Else
                            If valueAvailable Then
                                ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                                ' Throws an error if any other unescaped quotes.
                                value = ProcessSpecialCharactersInJsonStringValue(json, valueStartByte, valueEndByte, isValidatedString)
                            Else
                                ' Code will convert this to null.
                                value = vbNullString
                                isValidatedString = False
                            End If
                            AddValueToObject childObject, value, isValueString:=isValidatedString, keyForValue:=keyForValue
                            cursorByte = nextDelimitingCursorByte
                        End If
                        If isParentJsonObject Then keyForValue = vbNullString
                        value = vbNullString
                    Else
                        'Some error idk should be handled already.
                    End If
                End If
                
                With childObject
                    ' If cursor byte has exceeded its set range; remove it from the stack and update child and parent.
                    If cursorByte >= .Item(cursorLimitKey) Then
                        cursorByte = .Item(nextDelimitingParentCursorKey)
                        .Remove nextDelimitingParentCursorKey
                        .Remove cursorLimitKey
                        With objectStack
                            ' Set equal to parent of the following child.
                            If .count > 2 Then Set parentObject = .Item(.count - 2)
                            ' Remove the child from the heiarchy as it is no longer needed.
                            If .count > 1 Then .Remove .count
                            
                            If .count >= 2 Then
                                Set childObject = .Item(.count)
                            ElseIf .count = 1 Then
                                Set childObject = .Item(1)
                            End If
                        End With
                    End If
                End With
            Loop
            cursorByte = parentObject.Item(nextDelimitingParentCursorKey)
        Loop
        
        Set childObject = Nothing

        With parentObject
            .Remove cursorLimitKey
            .Remove nextDelimitingParentCursorKey
        End With
        
        Set ParseJSON = parentObject
        
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Bytes are wrong."
    End If
    
    Exit Function
CATCH_FAILURE:
    PropagateError Err, "ParseJSON"
End Function

Private Sub GetKeyAndValueIndexes(json() As Byte, currentTokenIndex&, forwardDelimiterIndex&, valueStartByte&, valueEndByte&, valueAvailable As Boolean, keyForObject$)

    Dim keyValueDelimiterTokenIndex&, isKeyWithinQuotes As Boolean, keyStartByte&, KeyEndByte&, includeForwardDelimiterInValue As Boolean
    
    Const quote& = 34, colon$ = ":"
    Select Case json(forwardDelimiterIndex)
        Case 93, 125
            includeForwardDelimiterInValue = True
        Case 44
            includeForwardDelimiterInValue = False
        Case Else
            ' Big error
    End Select
    
    'If non whitespace characters exist then record their location; otherwise throw an error.
    If TryGetWhitespaceCleanedEnclosingBytes(json, currentTokenIndex + BYTE_OFFSET, forwardDelimiterIndex + IIf(includeForwardDelimiterInValue, 0, -BYTE_OFFSET), keyStartByte, valueEndByte) Then
        ' Does key start with a qukote?
        If json(keyStartByte) = quote Then

            keyValueDelimiterTokenIndex = LocationOfCharacterNotBetweenTextTokens(json, colon, keyStartByte, valueEndByte)
            ' Is there a colon outside of quotes?
            If keyValueDelimiterTokenIndex > 0 Then
                ' Find boundaries of key using location of key-value delimiter.
                Call TryGetWhitespaceCleanedEnclosingBytes(json, keyStartByte, keyValueDelimiterTokenIndex - BYTE_OFFSET, keyStartByte, KeyEndByte)
                ' Find boundaries of value.
                valueAvailable = TryGetWhitespaceCleanedEnclosingBytes(json, keyValueDelimiterTokenIndex + BYTE_OFFSET, valueEndByte - (BYTE_OFFSET - 1), valueStartByte, valueEndByte)
                
                If (KeyEndByte < valueStartByte) Then
                    ' Will raise error if invalid key.
                    keyForObject = ProcessSpecialCharactersInJsonStringValue(json, keyStartByte, KeyEndByte, isKeyWithinQuotes)
                    If Not isKeyWithinQuotes Then Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Key must be enclosed in quotes." & CStr(GetArraySubset(json, keyStartByte, KeyEndByte))
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Key overlaps with value." & CStr(GetArraySubset(json, keyStartByte, valueEndByte))
                End If
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Key-Value delimiter not found." & CStr(GetArraySubset(json, keyStartByte, valueEndByte))
            End If
        Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Key doesn't start with a quote: " & CStr(GetArraySubset(json, keyStartByte, valueEndByte))
        End If
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Expected Key-Value pair not found."
    End If
    
    Exit Sub
PropagateError:
    PropagateError Err, "GetKeyAndValueIndexes"
End Sub

Private Sub AddValueToObject(parentObject As Object, Optional value$, Optional isValueString As Boolean, Optional newChild As Object, Optional keyForValue$, Optional addingNumber As Boolean, Optional numberToAdd As Long)
    
    Dim isParentJsonObject As Boolean: isParentJsonObject = (TypeName(parentObject) <> "Collection")

    On Error GoTo PropagateError
    
    If newChild Is Nothing Then
        If addingNumber Then
            ' This is just used for tracking byte locations.
            If isParentJsonObject Then
                parentObject.Add keyForValue, numberToAdd
            Else
                parentObject.Add numberToAdd, keyForValue
            End If
        ElseIf isValueString Then
            ' String value.
            If isParentJsonObject Then
                parentObject.Add keyForValue, value
            Else
                parentObject.Add value
            End If
        ElseIf IsNumeric(value) Then
            If isParentJsonObject Then
                parentObject.Add keyForValue, IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
            Else
                parentObject.Add IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
            End If
        Else
            Select Case value
                Case "true", "false"
                    If isParentJsonObject Then
                        parentObject.Add keyForValue, CBool(value)
                    Else
                        parentObject.Add CBool(value)
                    End If
                Case "null", vbNullString
                    If isParentJsonObject Then
                        parentObject.Add keyForValue, Null
                    Else
                        parentObject.Add Null
                    End If
                Case Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid format: Value isn't numeric and failed string literal checks. Ensure it's enclosed in unescaped quotes: " & value
            End Select
        End If
    Else
        If isParentJsonObject Then
            parentObject.Add keyForValue, newChild
        Else
            parentObject.Add newChild
        End If
    End If

    Exit Sub
PropagateError:
    PropagateError Err, "AddValueToObject"
End Sub

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
    
    On Error GoTo Propagate
    ' These are tokens that denote the start of a new value, or the start/end of a new object or array.
    Const comma$ = ",", leftSquare$ = "[", rightSquare$ = "]", leftCurly$ = "{", rightCurly$ = "}"
    
    If (maxTokenIndex - LBound(json) - 2) Mod 2 <> 0 And BYTE_OFFSET = 2 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, description:="maxTokenIndex is even when it should be odd."
    End If
    
    Select Case json(currentTokenIndex)
        Case 93, 125
            ' If ] or }
            currentTokenIndex = currentTokenIndex + BYTE_OFFSET
    End Select
    
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
            tempIndexBaseOne = previousClose ' + BYTE_OFFSET
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
        Case Is <= 0
            ' Determine the smallest value between } and ] if available.
            nextCurlyClose = InStrB(tempIndexBaseOne, json, rightCurly)
            nextSquareClose = InStrB(tempIndexBaseOne, json, rightSquare)

            If nextCurlyClose > 0 Then nextTokenDelimiterIndex = nextCurlyClose
            If nextSquareClose > 0 And nextSquareClose < nextTokenDelimiterIndex Then nextTokenDelimiterIndex = nextSquareClose
            
            If nextTokenDelimiterIndex > 0 Then nextTokenDelimiterIndex = nextTokenDelimiterIndex + lboundMinusOne
            If maxTokenIndex < nextTokenDelimiterIndex Then nextTokenDelimiterIndex = maxTokenIndex

        Case currentTokenIndex, Is > maxTokenIndex
            nextTokenDelimiterIndex = maxTokenIndex
    End Select
    
    Select Case json(nextTokenDelimiterIndex)
        Case 44, 125, 93
            ',}] Token is valid.
        Case Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid JSON format. Calculated token index doesn't reference a value delimiter: " & Chr(json(nextTokenDelimiterIndex))
    End Select
    
    GetNextTokenDelimiterIndex = nextTokenDelimiterIndex
    Exit Function
Propagate:
    PropagateError Err, "GetNextTokenDelimiterIndex"
End Function

Private Function ProcessSpecialCharactersInJsonStringValue(ByRef json() As Byte, ByRef startByte&, ByRef endByte&, Optional ByRef isRangeQuoteEnclosed As Boolean) As Byte()
'========================================================================================================================================
'Summary: Converts special characters that are lead by a \.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes backslash
'========================================================================================================================================
    Dim outputBytes() As Byte, iMainLoop&, outputIndex&, ignoreSpecialCharacter As Boolean, lboundMinusOne&, isDelimiterEnclosed As Boolean, _
    groupedBackslashCount&, countOfBackslashToKeep&, tempCount&, uniCode(8) As Byte, processBytes As Boolean, specialChars(6) As String, unescapedQuoteCount As Long
    
    Const quote As Byte = 34, backslash As Byte = 92

    specialChars(0) = Chr(quote): specialChars(1) = Chr(backslash)
    specialChars(2) = vbCr: specialChars(3) = vbTab
    specialChars(4) = vbLf: specialChars(5) = VBA.vbBack
    specialChars(6) = VBA.vbFormFeed

    '(LBound(json) - 1) will adjust for non 1 based arrays.
    lboundMinusOne = LBound(json) - 1
    ' Determine if there are any special characters in the byte range.
    For iMainLoop = LBound(specialChars) To UBound(specialChars) Step BYTE_OFFSET
        tempCount = InStrB(startByte - lboundMinusOne, json, specialChars(iMainLoop))
        ' If within bounds set processBytes equal to true and exit the for loop.
        If tempCount <= (endByte - lboundMinusOne) And tempCount > 0 Then
            processBytes = True: Exit For
        End If
    Next iMainLoop

    On Error GoTo PropagateError
    
    If processBytes Then
    
        ReDim outputBytes(startByte To endByte)
        'outputIndex is incremented by 2 whenever a value is added so initialize at -2.
        outputIndex = (-1 * BYTE_OFFSET) + startByte
        
        For iMainLoop = startByte To endByte Step BYTE_OFFSET
        
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
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                        
                    ElseIf Not ignoreSpecialCharacter Then
                        On Error GoTo 0
                        Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: Unescaped backslash at end of string: '\'"
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
      
        isRangeQuoteEnclosed = (Asc(outputBytes) = quote And outputBytes(outputIndex) = quote)
        isDelimiterEnclosed = (Asc(outputBytes) = 91 And outputBytes(outputIndex) = 93)
        If Not isDelimiterEnclosed Then isDelimiterEnclosed = (Asc(outputBytes) = 123 And outputBytes(outputIndex) = 125)
        
        ' If there are 2 unescaped quotes in the byte array and they occur at the end, then byte segement is a string.
        If unescapedQuoteCount = 2 And isRangeQuoteEnclosed Then
            'Value is a string.
            outputBytes = GetArraySubset(outputBytes, LBound(outputBytes) + BYTE_OFFSET, outputIndex - 1)
            'isRangeQuoteEnclosed = True
        ElseIf unescapedQuoteCount >= 1 And Not isDelimiterEnclosed Then
            ' Ensure quality of string value. ie don't allow "apple to be excepted as a string.
            Err.Raise INVALID_JSON_FORMAT_ERROR, description:="Invalid Format: String value isn't enclosed in quotes: " & vbNewLine & vbNewLine & CStr(GetArraySubset(outputBytes, LBound(outputBytes), outputIndex + (BYTE_OFFSET - 1)))
        Else
            ' + 1 to encapsulate character byte.
            ReDim Preserve outputBytes(LBound(outputBytes) To outputIndex + BYTE_OFFSET - 1)
        End If
    Else
        outputBytes = GetArraySubset(json, startByte, endByte)
    End If

    ProcessSpecialCharactersInJsonStringValue = outputBytes
    Exit Function
PropagateError:
    PropagateError Err, "ProcessSpecialCharactersInJsonStringValue"
End Function

Private Function TryGetWhitespaceCleanedEnclosingBytes(json() As Byte, ByVal regionStartByte&, ByVal regionEndIndex&, leftByteStart&, rightByteEnd&) As Boolean
'========================================================================================================================
'   Summary:
'       Searches for non [ space,newlines,line feeds and carriage returns ] from both ends of [json()] if they exist.
'       leftByteStart and rightByteEnd will be set to the relevant non-white space bytes.
'============================================================================================================
    Dim cleaningLeft As Boolean, cleaningRight As Boolean, iCount As Byte, errorLeft As Boolean, _
    errorRight As Boolean, cleanOnLeft As Boolean, cleanOnRight As Boolean

    'Max index should be the non null portion of a string byte.
    cleaningLeft = True: cleaningRight = True

    Do
        If cleaningLeft Then
            On Error GoTo CatchLeft
            
            Select Case json(regionStartByte)
                Case 32, 9, 10, 13
                    regionStartByte = regionStartByte + BYTE_OFFSET
                    cleaningLeft = Not regionStartByte > regionEndIndex And regionStartByte <= regionEndIndex
                    cleanOnLeft = True
                Case Else
                    cleaningLeft = False
                    leftByteStart = regionStartByte
            End Select
            
        End If
Try_Right:
        If cleaningRight Then
            On Error GoTo CatchRight
            Select Case json(regionEndIndex)
                Case 32, 9, 10, 13 '12
                    regionEndIndex = regionEndIndex - BYTE_OFFSET
                    cleaningRight = Not regionEndIndex < regionStartByte And regionEndIndex >= regionStartByte
                    cleanOnRight = True
                Case Else
                    cleaningRight = False
                    rightByteEnd = regionEndIndex + (BYTE_OFFSET - 1)
            End Select
        End If
NextLoop:
    Loop While cleaningLeft Or cleaningRight

    If (errorLeft And errorRight) Or (regionStartByte > regionEndIndex) Then
        'Byte sequence is just whitespace characters
        TryGetWhitespaceCleanedEnclosingBytes = False
    Else
        TryGetWhitespaceCleanedEnclosingBytes = True
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

Private Function GetArraySubset(query() As Byte, regionStartByte&, regionEndByte&) As Byte()

    Dim outputA() As Byte, iCount&
    ' Check to make sure section to copy has enough bytes. 2 bytes per character. Ex: 3char string  = 6 bytes
    If (regionEndByte - regionStartByte) Mod 2 = 0 And BYTE_OFFSET = 2 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, "GetArraySubset", "Attempted to create a string from an uneven length byte array."
    End If
    
    ReDim outputA(regionStartByte To regionEndByte)
    
    For iCount = regionStartByte To regionEndByte Step BYTE_OFFSET
        outputA(iCount) = query(iCount)
    Next iCount
    GetArraySubset = outputA
    
End Function

Private Function LocationOfCharacterNotBetweenTextTokens(ByRef query() As Byte, character$, startIndex&, maxOuterLoopTokenIndex&) As Long
'===============================================================================================================
' Summary
'   Search for the next [character] not contained within quotes.
' Remarks:
'   [character] within escaped quotes will be ignored.
'===============================================================================================================
    Dim characterLocationBaseOne&, leftQuoteIndexBaseOne&, loopCount&, _
    rightQuoteIndexBaseOne&, adjustLeftQuote As Long, isNotEscapedQuote As Boolean, locationOfQuoteBaseOne&, updateCharacterlocation As Boolean
    
    Const quote$ = """", backSlashCode As Byte = 92
    On Error GoTo Return_Default
    
    Dim lboundMinusOne&: lboundMinusOne = LBound(query) - 1
    ' Assign to index base 1.
    rightQuoteIndexBaseOne = startIndex - lboundMinusOne - BYTE_OFFSET
    updateCharacterlocation = True

    Do
        loopCount = loopCount + 1
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

    Loop Until (characterLocationBaseOne < leftQuoteIndexBaseOne And Not updateCharacterlocation) Or characterLocationBaseOne = 0 Or (characterLocationBaseOne > maxOuterLoopTokenIndex - lboundMinusOne) Or leftQuoteIndexBaseOne = 0 Or (loopCount = 1 And characterLocationBaseOne < leftQuoteIndexBaseOne)
    
    Select Case characterLocationBaseOne
        Case Is > maxOuterLoopTokenIndex, 0
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
''Author: Guido Witt-DÃ¶rring, https://stackoverflow.com/a/75787820/12287457
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
