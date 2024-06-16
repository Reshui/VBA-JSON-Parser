' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const INDEX_OFFSET As Byte = 1
#Const DebugActive = True

Option Explicit
Public Function ParseJsonString(ByRef json$) As Object
    
    'Dim jsonBytes() As Byte
    
    'jsonBytes = json
    'PreProcessData jsonBytes
    Set ParseJsonString = ParseJSON(json)
    
End Function

Public Function ParseJsonBytes(ByRef jsonBytes() As Byte) As Object
    'PreProcessData jsonBytes
    Set ParseJsonBytes = ParseJSON(CStr(jsonBytes))
End Function

Private Function ParseJSON(ByRef json$) As Object
''====================================================================================================================================================================================
'' Summary: Non-Recursive JSON byte parser.
'' Inputs:
''   jsonBytes: Byte array to parse.
'' Reference:
''   https://www.jsonBytes.org/jsonBytes-en.html
''====================================================================================================================================================================================
    Dim value$, isValueJsonObject As Boolean, isValueJsonArray As Boolean, isContainerJsonObject As Boolean, valueAvailable As Boolean

    Dim jsonLength&, keyForValue$, valueStartIndex&, valueEndIndex&, jsonContainer As Object, _
    isValidatedString As Boolean, heirarchyStack As New Collection, cursorLocation&, nextCursorLocation&, isKeyAvailable As Boolean

    Const rightCurly$ = "}", leftCurly$ = "{", leftSquare$ = "[", rightSquare$ = "]"
    
    Const cursorLimitKey$ = "MoshiM_UC@outlook.com"
    Const nextDelimitingParentCursorKey$ = "https://github.com/Reshui?tab=repositories"
    
    On Error GoTo CATCH_FAILURE
    
    json = Trim$(Replace$(Replace$(Replace$(json, vbTab, vbNullString), vbNewLine, vbNullString), vbLf, vbNullString))
    
    If Left$(json, 1) = leftCurly And Right$(json, 1) = rightCurly Then
        isValueJsonObject = True
    ElseIf Left$(json, 1) = leftSquare And Right$(json, 1) = rightSquare Then
        isValueJsonArray = True
    End If

    If isValueJsonObject Or isValueJsonArray Then

        cursorLocation = 1: jsonLength = Len(json)
        
        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
        Set jsonContainer = heirarchyStack(1)
        
        AddValueToObject jsonContainer, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=cursorLimitKey
        AddValueToObject jsonContainer, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=nextDelimitingParentCursorKey
        
        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
        
        Do While cursorLocation < jsonContainer.Item(cursorLimitKey)
            ' Finds the next token that encloses a value in conjunction with the current value of cursorLocation.
            nextCursorLocation = GetNextEnclosingDelimiterLocation(json, cursorLocation, jsonContainer.Item(cursorLimitKey))

            If nextCursorLocation > cursorLocation Then
            
                If isContainerJsonObject Then
                    ' Adjust by INDEX_OFFSET to ignore enclosing token delimiters.
                    GetValueRangeAndKey json, cursorLocation + INDEX_OFFSET, nextCursorLocation - INDEX_OFFSET, valueStartIndex, valueEndIndex, valueAvailable, keyForValue, isKeyAvailable
                End If
    
                If Not isContainerJsonObject Or (isKeyAvailable And isContainerJsonObject) Then
                                
                    If Not isContainerJsonObject Then
                        'Adjust by INDEX_OFFSET to exclude token delimiters.
                        valueAvailable = TryGetWhitespaceCleanedIndexes(json, cursorLocation + INDEX_OFFSET, nextCursorLocation - INDEX_OFFSET, valueStartIndex, valueEndIndex)
                    End If
                    'If value is enclosed in {} or [] then parse a new object or array.
                    isValueJsonObject = valueAvailable And (Mid$(json, valueStartIndex, 1) = leftCurly And Mid$(json, valueEndIndex, 1) = rightCurly)
                    isValueJsonArray = valueAvailable And (Mid$(json, valueStartIndex, 1) = leftSquare And Mid$(json, valueEndIndex, 1) = rightSquare)

                    If valueAvailable And (isValueJsonObject Or isValueJsonArray) Then

                        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
                        With heirarchyStack
                            ' Create heirarchy by nesting child within parent.
                            AddValueToObject jsonContainer, newChild:=.Item(.count), keyForValue:=keyForValue
                            Set jsonContainer = .Item(.count)
                        End With

                        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                        ' Add key to limit byte range for object.
                        AddValueToObject jsonContainer, addingNumber:=True, numberToAdd:=valueEndIndex, keyForValue:=cursorLimitKey
                        ' Byte to update cursorLocation to once array or object has been completed.
                        AddValueToObject jsonContainer, addingNumber:=True, numberToAdd:=nextCursorLocation, keyForValue:=nextDelimitingParentCursorKey
                        ' The token byte used to indicate the start of an object or array.
                        cursorLocation = valueStartIndex
                    Else
                        If valueAvailable Then
                            ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                            ' Throws an error if any other unescaped quotes.
                            value = ProcessSpecialCharactersInJsonStringValue(json, valueStartIndex, valueEndIndex, isValidatedString)
                        Else
                            ' Code will convert this to null in AddValueToObject() below.
                            value = vbNullString
                            isValidatedString = False
                        End If
                        AddValueToObject jsonContainer, value, isValueString:=isValidatedString, keyForValue:=keyForValue
                        If valueAvailable Then value = vbNullString
                        cursorLocation = nextCursorLocation
                    End If

                    If isKeyAvailable Then
                        keyForValue = vbNullString
                        isKeyAvailable = False
                    End If
                Else
                    'Should already be a raised error in GetValueRangeAndKey().
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="JSON parser failed. Key unavailable for object."
                End If
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="JSON parser failed. Current cursorLocation = next cursorLocation."
            End If
            
            With jsonContainer
                ' If cursorLocation byte has exceeded its set range; remove it from the stack.
                If cursorLocation >= .Item(cursorLimitKey) Then
                    cursorLocation = .Item(nextDelimitingParentCursorKey)
                    With heirarchyStack
                        ' Only update if not Parent of all nodes.
                        If .count > 1 Then
                            With jsonContainer
                                .Remove cursorLimitKey
                                .Remove nextDelimitingParentCursorKey
                            End With
                            .Remove .count
                            Set jsonContainer = .Item(.count)
                            isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                        End If
                    End With
                End If
            End With
        Loop
        
        With jsonContainer
            .Remove cursorLimitKey
            .Remove nextDelimitingParentCursorKey
        End With
        Set ParseJSON = jsonContainer
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Bytes are wrong."
    End If
    
    Exit Function
CATCH_FAILURE:
    #If DebugActive Then
        Stop: Resume
    #Else
        PropagateError Err, "ParseJSON"
    #End If
End Function

Private Sub CreateObjectOnStack(stack As Collection, createJsonObject As Boolean)

    If createJsonObject Then
        #If Mac Then
            ' https://github.com/VBA-tools/VBA-Dictionary
            stack.Add New Dictionary
        #Else
            stack.Add CreateObject("Scripting.Dictionary")
        #End If
    Else
        stack.Add New Collection
    End If
  
End Sub

Private Function GetNextEnclosingDelimiterLocation(json$, cursorLocation&, maxIndex&) As Long
'===========================================================================================================
' Summary:
'   Returns the location of the next token that encapsulates a complete value based on [cursorLocation]
' Notes:
'   [cursorLocation] should be set to one of  ,[{ which will note that the text that follows is a value/array/object.
'===========================================================================================================
    Dim nextSquareCloseBinary&, nextSquareOpenBinary&, nextValidCommaTokenIndexBinary&, currentIndexBinary&, nextOpenBinary&, nextCloseBinary&, minValues(1) As Long
    
    Dim maxIndexBinary&, nextCurlyCloseBinary&, nextCurlyOpenBinary&, previousCloseBinary&, tokenIndicator$(1), nextEnclosingTokenIndex&, escapeEnclosingArrayOrObject As Boolean

    On Error GoTo Propagate
    ' These are tokens that denote the start of a new value, or the start/end of a new object or array.
    Const comma$ = ",", leftSquare$ = "[", rightSquare$ = "]", leftCurly$ = "{", rightCurly$ = "}"
        
    ' Set to index of current token in array. Adjusted to be 1 based.
    currentIndexBinary = (cursorLocation * 2) - 1
    maxIndexBinary = (maxIndex * 2) - 1
    ' Start searching from the next byte.
    nextCurlyOpenBinary = InStrB(currentIndexBinary + INDEX_OFFSET, json, leftCurly)
    nextCurlyCloseBinary = InStrB(currentIndexBinary + INDEX_OFFSET, json, rightCurly)
    nextSquareOpenBinary = InStrB(currentIndexBinary + INDEX_OFFSET, json, leftSquare)
    nextSquareCloseBinary = InStrB(currentIndexBinary + INDEX_OFFSET, json, rightSquare)
    nextValidCommaTokenIndexBinary = BinaryLocationOfCharacterNotWithinQuotes(json, comma, cursorLocation + INDEX_OFFSET, maxIndex)

    If nextValidCommaTokenIndexBinary > 0 Then nextEnclosingTokenIndex = (nextValidCommaTokenIndexBinary \ 2) + 1

    If nextValidCommaTokenIndexBinary > 0 Then
        '======================================================================================================================================
        ' Determine which of [ or { is greater than cursorLocation AND appears first.
        ' If nextValidCommaTokenIndexBinary is greater than that byte,
        ' then determine if the comma is encapsulated. If it is, then find the next comma that isn't encapsulated to find the end of the value.
        '======================================================================================================================================
        If IsValueInRange(nextCurlyOpenBinary, currentIndexBinary + INDEX_OFFSET, nextValidCommaTokenIndexBinary - INDEX_OFFSET) And (nextCurlyOpenBinary < nextSquareOpenBinary Or nextSquareOpenBinary = 0) Then
            
            tokenIndicator(0) = leftCurly: tokenIndicator(1) = rightCurly
            nextOpenBinary = nextCurlyOpenBinary: nextCloseBinary = nextCurlyCloseBinary
            escapeEnclosingArrayOrObject = True
                  
        ElseIf IsValueInRange(nextSquareOpenBinary, currentIndexBinary + INDEX_OFFSET, nextValidCommaTokenIndexBinary - INDEX_OFFSET) And (nextSquareOpenBinary < nextCurlyOpenBinary Or nextCurlyOpenBinary = 0) Then
            
            tokenIndicator(0) = leftSquare: tokenIndicator(1) = rightSquare
            nextOpenBinary = nextSquareOpenBinary: nextCloseBinary = nextSquareCloseBinary
            escapeEnclosingArrayOrObject = True
            
        End If
    
        If escapeEnclosingArrayOrObject Then
            'Logic: Initialize previousCloseBinary to the next token End character found after the current position in the byte array.
            'Update the next token end and open position.
            'If open greater than the previous close or next open = 0 then you have exited the object/array that triggered the loop.
            Do
                previousCloseBinary = nextCloseBinary
                nextCloseBinary = InStrB(previousCloseBinary + INDEX_OFFSET, json, tokenIndicator(1))
                nextOpenBinary = InStrB(nextOpenBinary + INDEX_OFFSET, json, tokenIndicator(0))
            Loop Until nextOpenBinary > previousCloseBinary Or nextOpenBinary = 0 Or (previousCloseBinary >= maxIndexBinary)
            
            ' Check if still within bounds.
            If previousCloseBinary < maxIndexBinary Then
                                
                ' Find min not equal to 0 of the following.
                nextEnclosingTokenIndex = BinaryLocationOfCharacterNotWithinQuotes(json, comma, (previousCloseBinary \ 2) + 1, maxIndex)
                minValues(0) = InStrB(previousCloseBinary + INDEX_OFFSET, json, rightCurly)
                minValues(1) = InStrB(previousCloseBinary + INDEX_OFFSET, json, rightSquare)
                
                For nextCloseBinary = 0 To 1
                    If IsValueInRange(minValues(nextCloseBinary), 1, nextEnclosingTokenIndex) Or nextEnclosingTokenIndex <= 0 Then nextEnclosingTokenIndex = minValues(nextCloseBinary)
                Next nextCloseBinary
                If nextEnclosingTokenIndex > 0 Then nextEnclosingTokenIndex = (nextEnclosingTokenIndex \ 2) + 1
            Else
                nextEnclosingTokenIndex = maxIndex
            End If
        End If
    End If

    Select Case nextEnclosingTokenIndex
        Case Is <= 0, cursorLocation, Is > maxIndex
            nextEnclosingTokenIndex = maxIndex
    End Select
    
    If InStrB(1, ",]}", Mid$(json, nextEnclosingTokenIndex, 1)) = 0 Then
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid JSON format. Calculated token index doesn't reference a value delimiter: " & Mid$(json, nextEnclosingTokenIndex, 1)
    End If
    
    GetNextEnclosingDelimiterLocation = nextEnclosingTokenIndex
    Exit Function
Propagate:
    #If DebugActive Then
        Stop: Resume
    #Else
        PropagateError Err, "GetNextEnclosingDelimiterLocation"
    #End If

End Function

Private Function IsValueInRange(numberToCheck&, lowerBound&, upperBound&) As Boolean
'Summary Checks if numberToCheck is contained within a given range.
    IsValueInRange = numberToCheck >= lowerBound And numberToCheck <= upperBound
End Function

Private Sub GetValueRangeAndKey(ByRef json$, ByRef minimumIndex&, ByRef maxIndex&, ByRef valueStartLocation&, ByRef valueEndLocation&, ByRef valueAvailable As Boolean, ByRef keyForObject$, ByRef isKeyAvailalble As Boolean)

    Dim keyValueDelimiterIndex&, isKeyWithinQuotes As Boolean, keyStartLocation&, keyEndLocation&
    
    Const quote$ = """", colon$ = ":"
    On Error GoTo PropagateError
    'If non whitespace characters exist then record their location; otherwise throw an error.
    If TryGetWhitespaceCleanedIndexes(json, minimumIndex, maxIndex, keyStartLocation, valueEndLocation) Then
        ' Does key start with a qukote?
        If Mid$(json, keyStartLocation, 1) = quote Then

            keyValueDelimiterIndex = (BinaryLocationOfCharacterNotWithinQuotes(json, colon, keyStartLocation, valueEndLocation) \ 2) + 1
            ' Is there a colon outside of quotes?
            If keyValueDelimiterIndex > 0 Then
                ' Find boundaries of key using location of key-value delimiter.
                Call TryGetWhitespaceCleanedIndexes(json, keyStartLocation, keyValueDelimiterIndex - INDEX_OFFSET, keyStartLocation, keyEndLocation)
                ' Find boundaries of value.
                valueAvailable = TryGetWhitespaceCleanedIndexes(json, keyValueDelimiterIndex + INDEX_OFFSET, valueEndLocation, valueStartLocation, valueEndLocation)
                
                If (keyEndLocation < valueStartLocation) Then
                    ' Will raise error if invalid key.
                    keyForObject = ProcessSpecialCharactersInJsonStringValue(json, keyStartLocation, keyEndLocation, isKeyWithinQuotes)
                    If Not isKeyWithinQuotes Then Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key must be enclosed in quotes." & Mid$(json, keyStartLocation, keyEndLocation - keyStartLocation + 1)
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key overlaps with value." & Mid$(json, keyStartLocation, keyEndLocation - valueEndLocation + 1)
                End If
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key-Value delimiter not found." & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
            End If
        Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key doesn't start with a quote: " & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
        End If
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Expected Key-Value pair not found."
    End If

    isKeyAvailalble = True
    Exit Sub
PropagateError:
    PropagateError Err, "GetValueRangeAndKey"
End Sub

Private Sub AddValueToObject(containerObject As Object, Optional value$, Optional isValueString As Boolean, Optional newChild As Object, Optional keyForValue$, Optional addingNumber As Boolean, Optional numberToAdd As Long)
    
    Dim isContainerJsonObject As Boolean: isContainerJsonObject = Not TypeOf containerObject Is Collection

    On Error GoTo PropagateError
    
    If newChild Is Nothing Then
        If addingNumber Then
            ' This is just used for tracking byte locations.
            If isContainerJsonObject Then
                containerObject.Add keyForValue, numberToAdd
            Else
                containerObject.Add numberToAdd, keyForValue
            End If
        ElseIf isValueString Then
            ' String value.
            If isContainerJsonObject Then
                containerObject.Add keyForValue, value
            Else
                containerObject.Add value
            End If
        ElseIf IsNumeric(value) Then
            If isContainerJsonObject Then
                containerObject.Add keyForValue, IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
            Else
                containerObject.Add IIf(InStrB(1, value, ".") > 0, CDbl(value), CLng(value))
            End If
        Else
            Select Case value
                Case "true", "false"
                    If isContainerJsonObject Then
                        containerObject.Add keyForValue, CBool(value)
                    Else
                        containerObject.Add CBool(value)
                    End If
                Case "null", vbNullString
                    If isContainerJsonObject Then
                        containerObject.Add keyForValue, Null
                    Else
                        containerObject.Add Null
                    End If
                Case Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid format: Value isn't numeric and failed string literal checks. Ensure it's enclosed in unescaped quotes: " & value
            End Select
        End If
    Else
        If isContainerJsonObject Then
            containerObject.Add keyForValue, newChild
        Else
            containerObject.Add newChild
        End If
    End If

    Exit Sub
PropagateError:
    PropagateError Err, "AddValueToObject"
End Sub

Private Function ProcessSpecialCharactersInJsonStringValue(ByRef json$, ByRef startIndex&, ByRef lastIndex&, Optional ByRef isRangeQuoteEnclosed As Boolean) As String
'========================================================================================================================================
'Summary: Converts special characters that are lead by a \.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes backslash
'========================================================================================================================================
    Dim outputString$, iMainLoop&, outputIndex&, ignoreSpecialCharacter As Boolean, isDelimiterEnclosed As Boolean, currentCharacter$, _
    groupedBackslashCount&, countOfBackslashToKeep&, processString As Boolean, specialChars(6) As String, unescapedQuoteCount As Long
    
    Const quote$ = """", backslash As String = "\"

    specialChars(0) = quote: specialChars(1) = backslash
    specialChars(2) = vbCr: specialChars(3) = vbTab
    specialChars(4) = vbLf: specialChars(5) = VBA.vbBack
    specialChars(6) = VBA.vbFormFeed

    ' Determine if there are any special characters in the byte range.
    For iMainLoop = LBound(specialChars) To UBound(specialChars) Step INDEX_OFFSET
        ' If within bounds set processString equal to true and exit the for loop.
        If IsValueInRange(InStrB((startIndex * 2) - 1, json, specialChars(iMainLoop)), 1, (lastIndex * 2) - 1) Then
            processString = True: Exit For
        End If
    Next iMainLoop

    On Error GoTo PropagateError
    
    If processString Then

        outputIndex = 0
        For iMainLoop = startIndex To lastIndex Step INDEX_OFFSET

            currentCharacter = Mid$(json, iMainLoop, 1)
            Select Case currentCharacter
                Case backslash
                    ' \ encountered.
                    groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        iMainLoop = iMainLoop + INDEX_OFFSET
                        If iMainLoop > Len(json) Then Exit Do
                    Loop Until Mid$(json, iMainLoop, 1) <> backslash
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    
                    If countOfBackslashToKeep > 0 Then
                        ' Write backslashes to output array.
                        For outputIndex = outputIndex + INDEX_OFFSET To outputIndex + (2 * countOfBackslashToKeep) Step INDEX_OFFSET
                            Mid$(json, outputIndex, 1) = backslash
                        Next outputIndex
                        ' For loops will exceed boundary on last loop so reverse by the value of step.
                        outputIndex = outputIndex - INDEX_OFFSET
                    End If
                    
                    ignoreSpecialCharacter = groupedBackslashCount Mod 2 = 0
                    
                    If iMainLoop <= Len(json) Then
                        currentCharacter = Mid$(json, iMainLoop, 1)
                        outputIndex = outputIndex + INDEX_OFFSET
                        
                        If ignoreSpecialCharacter Then
                            Mid$(json, outputIndex, 1) = currentCharacter
                        Else
                            Select Case currentCharacter
                                Case "b" '\b
                                    Mid$(json, outputIndex, 1) = VBA.vbBack  'backspace
                                Case "n" '\n
                                    Mid$(json, outputIndex, 1) = vbNewLine 'newline
                                Case "t" '\t
                                    Mid$(json, outputIndex, 1) = vbTab  'tab
                                Case "0"
                                Case quote, "'", "/" '\" or \/ or \0 or \'
                                    Mid$(json, outputIndex, 1) = currentCharacter
                                Case "r" '\r
                                    Mid$(json, outputIndex, 1) = vbCr 'carriage return
                                Case "f" '\f
                                    Mid$(json, outputIndex, 1) = vbFormFeed ' Form feed
                                Case "u" '\u
                                    ' See: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chr-function
                                    Mid$(json, outputIndex, 1) = ChrW(Val("&h" & Mid$(json, iMainLoop + 1, 4)))
                                    iMainLoop = iMainLoop + 4
                                Case Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                    ElseIf Not ignoreSpecialCharacter Then
                        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped backslash at end of string: '\'"
                    End If
                'Case 47
                    'Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped solidius: '/'"
                Case Else
                    outputIndex = outputIndex + INDEX_OFFSET
                    Mid$(json, outputIndex, 1) = currentCharacter
                    If currentCharacter = quote Then unescapedQuoteCount = unescapedQuoteCount + 1
            End Select
            
        Next iMainLoop
      
        isRangeQuoteEnclosed = Left$(json, 1) = quote And Mid$(json, outputIndex, 1) = quote
        
        If Not isRangeQuoteEnclosed Then
            If (Left$(json, 1) = "[" And Mid$(json, outputIndex, 1) = "]") Then
                isDelimiterEnclosed = True
            ElseIf (Left$(json, 1) = "{" And Mid$(json, outputIndex, 1) = "}") Then
                isDelimiterEnclosed = True
            End If
        End If
        
        ' If there are 2 unescaped quotes in the byte array and they occur at the end, then byte segement is a string.
        If (unescapedQuoteCount = 2 And isRangeQuoteEnclosed) Then
            outputString = Mid$(json, 2, outputIndex - 2)
        ElseIf unescapedQuoteCount >= 1 And Not isDelimiterEnclosed Then
            ' Ensure quality of string value. ie don't allow "apple to be excepted as a string.
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: String value isn't enclosed in quotes: " & Left$(json, outputIndex)
        Else
            outputString = Left$(json, outputIndex)
        End If
    Else
        isRangeQuoteEnclosed = False
        outputString = Mid$(json, startIndex, lastIndex - startIndex + 1)
    End If

    ProcessSpecialCharactersInJsonStringValue = outputString
    Exit Function
PropagateError:
    PropagateError Err, "ProcessSpecialCharactersInJsonStringValue"
End Function

Private Function TryGetWhitespaceCleanedIndexes(json As String, ByVal regionStartIndex&, ByVal regionEndIndex&, cleanedStartIndex&, cleanedEndIndex&) As Boolean
'========================================================================================================================
'   Summary:
'       Searches for non [ space,newlines,line feeds and carriage returns ] from both ends of [json()] if they exist.
'       cleanedStartIndex and cleanedEndIndex will be set to the relevant non-white space bytes.
'============================================================================================================
    Dim cleaningLeft As Boolean, cleaningRight As Boolean, errorLeft As Boolean, errorRight As Boolean

    'Max index should be the non null portion of a string byte.
    cleaningLeft = True: cleaningRight = True

    Do
        If cleaningLeft Then
            On Error GoTo CatchLeft
            
            Select Case Mid$(json, regionStartIndex, 1)
                Case " " '32, 9, 10, 13
                    regionStartIndex = regionStartIndex + INDEX_OFFSET
                    cleaningLeft = Not regionStartIndex > regionEndIndex And regionStartIndex <= regionEndIndex
                Case Else
                    cleaningLeft = False
                    cleanedStartIndex = regionStartIndex
            End Select
            
        End If
Try_Right:
        If cleaningRight Then
            On Error GoTo CatchRight
            Select Case Mid$(json, regionEndIndex, 1)
                Case " " '32, 9, 10, 13 '12
                    regionEndIndex = regionEndIndex - INDEX_OFFSET
                    cleaningRight = Not regionEndIndex < regionStartIndex And regionEndIndex >= regionStartIndex
                Case Else
                    cleaningRight = False
                    cleanedEndIndex = regionEndIndex
            End Select
        End If
NextLoop:
    Loop While cleaningLeft Or cleaningRight

    If (errorLeft And errorRight) Or (regionStartIndex > regionEndIndex) Then
        'Byte sequence is just whitespace characters
        TryGetWhitespaceCleanedIndexes = False
    Else
        TryGetWhitespaceCleanedIndexes = True
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

Private Function BinaryLocationOfCharacterNotWithinQuotes(ByRef query$, ByRef character$, ByRef startIndex&, ByRef maxIndex&) As Long
'===============================================================================================================
' Summary
'   Search for the next [character] not contained within quotes.
' Remarks:
'   [character] within escaped quotes will be ignored.
'===============================================================================================================
    Dim characterIndexBinary&, leftQuoteIndexBinary&, loopCount&, rightQuoteIndexBinary&, _
    adjustLeftQuote As Long, isNotEscapedQuote As Boolean, quoteIndexBinary&, updateCharacterlocation As Boolean, maxIndexBinary&
    
    Const quote$ = """", backslash$ = "\", BINARY_INDEX_OFFSET As Byte = 2
    On Error GoTo Return_Default
    
    maxIndexBinary = (maxIndex * 2) - 1
    rightQuoteIndexBinary = ((startIndex * 2) - 1) - BINARY_INDEX_OFFSET
    
    updateCharacterlocation = True
    
    Do
        If updateCharacterlocation Then
            characterIndexBinary = InStrB(rightQuoteIndexBinary + BINARY_INDEX_OFFSET, query, character)
            If characterIndexBinary = 0 Then Exit Do
        End If
        
        loopCount = loopCount + 1
        ' Start searching from the location of right quote: wiil be incremented before searching.
        leftQuoteIndexBinary = rightQuoteIndexBinary

        For adjustLeftQuote = -1 To 0

            quoteIndexBinary = leftQuoteIndexBinary
            Do
                quoteIndexBinary = InStrB(quoteIndexBinary + BINARY_INDEX_OFFSET, query, quote)
                If Not quoteIndexBinary = 0 Then
                    ' Set to TRUE if previous character in the string isn't a backslash.
                    isNotEscapedQuote = Mid$(query, ((quoteIndexBinary \ 2) + 1) - INDEX_OFFSET, 1) <> backslash
                End If
            Loop Until isNotEscapedQuote Or quoteIndexBinary = 0 Or (loopCount = 1 And adjustLeftQuote And characterIndexBinary < quoteIndexBinary And isNotEscapedQuote)
            
            If adjustLeftQuote Then
                leftQuoteIndexBinary = quoteIndexBinary
            Else
                rightQuoteIndexBinary = quoteIndexBinary
            End If
            
        Next adjustLeftQuote
        ' Update if sandwhiched between quotes.
        updateCharacterlocation = (characterIndexBinary > leftQuoteIndexBinary And characterIndexBinary < rightQuoteIndexBinary)

    Loop Until (characterIndexBinary < leftQuoteIndexBinary And Not updateCharacterlocation) Or characterIndexBinary = 0 Or (characterIndexBinary > maxIndexBinary) Or leftQuoteIndexBinary = 0 Or (loopCount = 1 And characterIndexBinary < leftQuoteIndexBinary)
    
    If characterIndexBinary > maxIndexBinary Or updateCharacterlocation Then
        BinaryLocationOfCharacterNotWithinQuotes = 0
    ElseIf Not updateCharacterlocation Then
        BinaryLocationOfCharacterNotWithinQuotes = characterIndexBinary
    End If
    
    Exit Function
Return_Default:
    BinaryLocationOfCharacterNotWithinQuotes = 0
End Function

Private Sub PropagateError(e As ErrObject, procedureName$, Optional moreDetails$ = vbNullString)
    Dim firstPropagation As Boolean, sourceParts$()
    
    Const delim$ = ": "
    With e
        AddParentToErrSource e, procedureName
        If LenB(moreDetails) > 0 Then .Description = moreDetails & vbNewLine & .Description
        .Raise .number, .source, .Description
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
End Sub

Public Sub Unit_Test()
    Dim success As Boolean
    With Me
        '.ParseJsonString "{""apple"":""tree"",""inner object"":1,""array"":[1,2,3,4]}" ', True
        .ParseJsonString HTTPGet("http://worldtimeapi.org/api/timezone/America/New_York", success) ', True
        '.ParseJsonString HTTPGet("https://publicreporting.cftc.gov/resource/jun7-fc8e.json?$select=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name&$group=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name&$limit=5000", success)', True
    End With
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

