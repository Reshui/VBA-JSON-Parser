' JSON parser for VBA.
'
' @author: MoshiM_UC@outlook.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'To DO hex to utf-16 for mac users since chrw doesn't work.
' https://stackoverflow.com/a/75787820/12287457; https://github.com/guwidoe/VBA-StringTools

Option Explicit
Private Const INVALID_JSON_FORMAT_ERROR& = vbObjectError + 547
Private Const INDEX_OFFSET As Byte = 1
Private Const RightCurly$ = "}", LeftCurly$ = "{", LeftSquare$ = "[", RightSquare$ = "]", _
QuoteCharacter$ = """", BackslashCharacter$ = "\", SpaceCharacter$ = " ", ColonCharacter$ = ":", Comma$ = ",", BINARY_OFFSET As Byte = 2

#Const DebugActive = False
Public Function ParseJsonString(ByRef json$) As Object
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
''   https://www.json.org/json-en.html
''====================================================================================================================================================================================
    Dim value$, isValueJsonObject As Boolean, isValueJsonArray As Boolean, isContainerJsonObject As Boolean, valueAvailable As Boolean

    Dim jsonLength&, keyForValue$, valueStartIndex&, valueEndIndex&, jsonContainer As Object, _
    isValidatedString As Boolean, heirarchyStack As New Collection, cursorLocation&, nextCursorLocation&, isKeyAvailable As Boolean
    
    Const cursorLimit$ = "MoshiM_UC@outlook.com"
    Const cursorLocationAfterCompletion$ = "https://github.com/Reshui?tab=repositories"
    
    On Error GoTo CATCH_FAILURE
    ' This is basically a trim that also works on vbLf,vbTab,vbCr and doesn't return a new string just start and end locations.
    If TryGetWhitespaceCleanedIndexes(json, 1, Len(json), cursorLocation, jsonLength) Then
        If Mid$(json, cursorLocation, 1) = LeftCurly And Mid$(json, jsonLength, 1) = RightCurly Then
            isValueJsonObject = True
        ElseIf Mid$(json, cursorLocation, 1) = LeftSquare And Mid$(json, jsonLength, 1) = RightSquare Then
            isValueJsonArray = True
        End If
    End If
    
    If isValueJsonObject Or isValueJsonArray Then
        
        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
        Set jsonContainer = heirarchyStack(1)
        
        AddValueToObject jsonContainer, isValueJsonObject, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=cursorLimit
        AddValueToObject jsonContainer, isValueJsonObject, addingNumber:=True, numberToAdd:=jsonLength, keyForValue:=cursorLocationAfterCompletion
        
        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
        
        Do While cursorLocation < jsonContainer.Item(cursorLimit)
            ' Finds the next token that encloses a value in conjunction with the current value of cursorLocation.
            nextCursorLocation = GetNextEnclosingDelimiterLocation(json, cursorLocation, jsonContainer.Item(cursorLimit))

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
                    isValueJsonObject = valueAvailable And (Mid$(json, valueStartIndex, 1) = LeftCurly And Mid$(json, valueEndIndex, 1) = RightCurly)
                    isValueJsonArray = valueAvailable And (Mid$(json, valueStartIndex, 1) = LeftSquare And Mid$(json, valueEndIndex, 1) = RightSquare)

                    If valueAvailable And (isValueJsonObject Or isValueJsonArray) Then

                        CreateObjectOnStack heirarchyStack, createJsonObject:=isValueJsonObject
                        With heirarchyStack
                            ' Create heirarchy by nesting child within parent.
                            AddValueToObject jsonContainer, isContainerJsonObject, newChild:=.Item(.count), keyForValue:=keyForValue
                            Set jsonContainer = .Item(.count)
                        End With

                        isContainerJsonObject = Not TypeOf jsonContainer Is Collection
                        ' Add key to limit byte range for object.
                        AddValueToObject jsonContainer, isContainerJsonObject, addingNumber:=True, numberToAdd:=valueEndIndex, keyForValue:=cursorLimit
                        ' Byte to update cursorLocation to once array or object has been completed.
                        AddValueToObject jsonContainer, isContainerJsonObject, addingNumber:=True, numberToAdd:=nextCursorLocation, keyForValue:=cursorLocationAfterCompletion
                        ' The token byte used to indicate the start of an object or array.
                        cursorLocation = valueStartIndex
                    Else
                        If valueAvailable Then
                            ' This function will remove a pair of unescaped quotes at the beginning and end of string values.
                            ' Throws an error if any other unescaped quotes.
                            value = ProcessCharactersInRange(json, valueStartIndex, valueEndIndex, isValidatedString)
                        Else
                            ' Code will convert this to null in AddValueToObject() below.
                            value = vbNullString
                            isValidatedString = False
                        End If
                        AddValueToObject jsonContainer, isContainerJsonObject, value, isValueString:=isValidatedString, keyForValue:=keyForValue
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
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="JSON parser failed. Current cursorLocation >= next cursorLocation."
            End If
            
            With jsonContainer
                ' If cursorLocation has exceeded jsonContainer Cursor Limit; remove it from the stack.
                If cursorLocation >= .Item(cursorLimit) Then
                    cursorLocation = .Item(cursorLocationAfterCompletion)
                    With heirarchyStack
                        ' Only update if not Parent of all nodes.
                        If .count > 1 Then
                            With jsonContainer
                                .Remove cursorLimit
                                .Remove cursorLocationAfterCompletion
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
            .Remove cursorLimit
            .Remove cursorLocationAfterCompletion
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
'   Returns the location of the next token that encapsulates a complete value based on [cursorLocation].
'===========================================================================================================
    Dim nextSquareCloseBinary&, nextSquareOpenBinary&, nextValidCommaTokenIndexBinary&, currentIndexBinary&, nextOpenBinary&, nextCloseBinary&, minValues(1) As Long
    
    Dim maxIndexBinary&, nextCurlyCloseBinary&, nextCurlyOpenBinary&, previousCloseBinary&, tokenIndicator$(1), nextEnclosingTokenIndex&, escapeEnclosingArrayOrObject As Boolean

    On Error GoTo Propagate
    ' Convert to binary location.
    currentIndexBinary = (cursorLocation * BINARY_OFFSET) - 1
    maxIndexBinary = (maxIndex * BINARY_OFFSET) - 1
    ' Start searching from the next byte.
    nextCurlyOpenBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, LeftCurly)
    nextCurlyCloseBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, RightCurly)
    nextSquareOpenBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, LeftSquare)
    nextSquareCloseBinary = InStrB(currentIndexBinary + BINARY_OFFSET, json, RightSquare)
    nextValidCommaTokenIndexBinary = BinaryCharacterIndexNotWithinQuotes(json, Comma, cursorLocation + INDEX_OFFSET, maxIndex)

    If nextValidCommaTokenIndexBinary > 0 Then nextEnclosingTokenIndex = (nextValidCommaTokenIndexBinary \ BINARY_OFFSET) + 1

    If nextValidCommaTokenIndexBinary > 0 Then
        '======================================================================================================================================
        ' Determine which of [ or { is greater than cursorLocation AND appears first.
        ' If nextValidCommaTokenIndexBinary is greater than that byte,
        ' then determine if the Comma is encapsulated. If it is, then find the next Comma that isn't encapsulated to find the end of the value.
        '======================================================================================================================================
        If IsValueInRange(nextCurlyOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary - BINARY_OFFSET) And (nextCurlyOpenBinary < nextSquareOpenBinary Or nextSquareOpenBinary = 0) Then
            
            tokenIndicator(0) = LeftCurly: tokenIndicator(1) = RightCurly
            nextOpenBinary = nextCurlyOpenBinary: nextCloseBinary = nextCurlyCloseBinary
            escapeEnclosingArrayOrObject = True
                  
        ElseIf IsValueInRange(nextSquareOpenBinary, currentIndexBinary + BINARY_OFFSET, nextValidCommaTokenIndexBinary - BINARY_OFFSET) And (nextSquareOpenBinary < nextCurlyOpenBinary Or nextCurlyOpenBinary = 0) Then
            
            tokenIndicator(0) = LeftSquare: tokenIndicator(1) = RightSquare
            nextOpenBinary = nextSquareOpenBinary: nextCloseBinary = nextSquareCloseBinary
            escapeEnclosingArrayOrObject = True
            
        End If
    
        If escapeEnclosingArrayOrObject Then
            'Logic: Initialize previousCloseBinary to the next token End character found after the current position in the byte array.
            'Update the next token end and open position.
            'If open greater than the previous close or next open = 0 then you have exited the object/array that triggered the loop.
            Do
                previousCloseBinary = nextCloseBinary
                nextCloseBinary = InStrB(previousCloseBinary + BINARY_OFFSET, json, tokenIndicator(1))
                nextOpenBinary = InStrB(nextOpenBinary + BINARY_OFFSET, json, tokenIndicator(0))
            Loop Until nextOpenBinary > previousCloseBinary Or nextOpenBinary = 0 Or (previousCloseBinary >= maxIndexBinary)

            If previousCloseBinary < maxIndexBinary Then
                ' Find min not equal to 0 of the following.
                nextEnclosingTokenIndex = BinaryCharacterIndexNotWithinQuotes(json, Comma, (previousCloseBinary \ BINARY_OFFSET) + 1, maxIndex)
                minValues(0) = InStrB(previousCloseBinary + BINARY_OFFSET, json, RightCurly)
                minValues(1) = InStrB(previousCloseBinary + BINARY_OFFSET, json, RightSquare)
                
                For nextCloseBinary = 0 To 1
                    If IsValueInRange(minValues(nextCloseBinary), 1, nextEnclosingTokenIndex) Or nextEnclosingTokenIndex <= 0 Then nextEnclosingTokenIndex = minValues(nextCloseBinary)
                Next nextCloseBinary
                If nextEnclosingTokenIndex > 0 Then nextEnclosingTokenIndex = (nextEnclosingTokenIndex \ BINARY_OFFSET) + 1
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

Private Sub GetValueRangeAndKey(ByRef json$, ByRef minimumIndex&, ByRef maxIndex&, ByRef valueStartLocation&, ByRef valueEndLocation&, ByRef valueAvailable As Boolean, ByRef keyForValue$, ByRef isKeyAvailalble As Boolean)
'============================================================================================================================================================================================================================================
'Summary: Attempts to retrieve a key within the bounds of [minimumIndex] and [maxIndex]. If successful then the indexes for the start and end of the value will be stored in
'   [Byref: valueStartLocation] and [ByRef:valueEndLocation]
'Variables:
'   [keyForValue] Used to store keys if found.
'============================================================================================================================================================================================================================================
    Dim keyValueDelimiterIndex&, keyStartLocation&, keyEndLocation&
    
    On Error GoTo PropagateError
    'If non whitespace characters exist then record their location; otherwise throw an error.
    If TryGetWhitespaceCleanedIndexes(json, minimumIndex, maxIndex, keyStartLocation, valueEndLocation) Then
        ' Does key start with a qukote?
        If Mid$(json, keyStartLocation, 1) = QuoteCharacter Then

            keyValueDelimiterIndex = (BinaryCharacterIndexNotWithinQuotes(json, ColonCharacter, keyStartLocation, valueEndLocation) \ BINARY_OFFSET) + 1
            ' Is there a ColonCharacter outside of quotes?
            If keyValueDelimiterIndex > 0 Then
                ' Find boundaries of key using location of key-value delimiter.
                Call TryGetWhitespaceCleanedIndexes(json, keyStartLocation, keyValueDelimiterIndex - INDEX_OFFSET, keyStartLocation, keyEndLocation)
                ' Find boundaries of value.
                valueAvailable = TryGetWhitespaceCleanedIndexes(json, keyValueDelimiterIndex + INDEX_OFFSET, valueEndLocation, valueStartLocation, valueEndLocation)
                
                If (keyEndLocation < valueStartLocation) Then
                    ' Will raise error if invalid key.
                    keyForValue = ProcessCharactersInRange(json, keyStartLocation, keyEndLocation, isKeyAvailalble)
                    If Not isKeyAvailalble Then Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key must be enclosed in quotes." & Mid$(json, keyStartLocation, keyEndLocation - keyStartLocation + 1)
                Else
                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key overlaps with value." & Mid$(json, keyStartLocation, keyEndLocation - valueEndLocation + 1)
                End If
            Else
                Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key-Value delimiter not found." & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
            End If
        Else
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Key doesn't start with a quote or couldn't find valid enclosing token: " & Mid$(json, keyStartLocation, keyStartLocation - valueEndLocation + 1)
        End If
    Else
        Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Expected Key-Value pair not found."
    End If

    isKeyAvailalble = True
    Exit Sub
PropagateError:
    PropagateError Err, "GetValueRangeAndKey"
End Sub

Private Sub AddValueToObject(containerObject As Object, isContainerJsonObject As Boolean, Optional value$, Optional isValueString As Boolean, Optional newChild As Object, Optional keyForValue$, Optional addingNumber As Boolean, Optional numberToAdd As Long)

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

Private Function ProcessCharactersInRange(ByRef json$, ByRef startIndex&, ByRef lastIndex&, Optional ByRef isRangeQuoteEnclosed As Boolean) As String
'========================================================================================================================================
'Summary: Converts special characters that are lead by a \ and raises error where appropriate.
'   For example unescaped quotes that don't enclose a value or no special character following a \.
'Note:
' Front end of [json] will be used as a buffer.
' \n: new line; \t: tab; \b: backspace; \r: carriage return; \f: form feed; \{/0"}: same character if not escaped; \\: escapes BackslashCharacter
'========================================================================================================================================
    Dim outputString$, iLoop&, outputIndex&, ignoreSpecialCharacter As Boolean, isDelimiterEnclosed As Boolean, currentCharacter$, minValues&(6), _
    groupedBackslashCount&, countOfBackslashToKeep&, processString As Boolean, specialChars(6) As String, unescapedQuoteCount As Long

    specialChars(0) = QuoteCharacter: specialChars(1) = BackslashCharacter
    specialChars(2) = vbCr: specialChars(3) = vbTab
    specialChars(4) = vbLf: specialChars(5) = VBA.vbBack
    specialChars(6) = VBA.vbFormFeed

    ' Determine if there are any special characters in the range.
    For iLoop = LBound(specialChars) To UBound(specialChars)
        ' If within bounds set processString equal to true and exit the loop.
        If IsValueInRange(InStrB((startIndex * BINARY_OFFSET) - 1, json, specialChars(iLoop)), 1, (lastIndex * BINARY_OFFSET) - 1) Then
            processString = True: Exit For
        End If
    Next iLoop

    On Error GoTo PropagateError
    
    If processString Then

        Dim characterIndex&, regionStart&, writeToBuffer As Boolean, nonInclusiveRegionEnd&, newOutputIndex&, newRegionStart&

        regionStart = startIndex: characterIndex = startIndex - INDEX_OFFSET: outputIndex = 0
        
        'Use InstrB to determine the forward index of any special characters.
        Do While characterIndex < lastIndex
            ' Convert to binary.
            characterIndex = (characterIndex * BINARY_OFFSET) - 1 + BINARY_OFFSET
            ' Determine min special char > 0
            For iLoop = LBound(specialChars) To UBound(specialChars)
                If minValues(iLoop) < characterIndex Then minValues(iLoop) = InStrB(characterIndex, json, specialChars(iLoop))
                If IsValueInRange(minValues(iLoop), 1, characterIndex - BINARY_OFFSET) Or iLoop = LBound(specialChars) Then characterIndex = minValues(iLoop)
            Next iLoop
            ' Convert from binary with integer division.
            If characterIndex > 0 Then characterIndex = (characterIndex \ BINARY_OFFSET) + 1

            If characterIndex > lastIndex Or characterIndex = 0 Then
                characterIndex = lastIndex
            End If

            currentCharacter = Mid$(json, characterIndex, 1)

            Select Case currentCharacter

                Case QuoteCharacter
                    unescapedQuoteCount = unescapedQuoteCount + 1: writeToBuffer = True: nonInclusiveRegionEnd = characterIndex + 1
                    newRegionStart = characterIndex + INDEX_OFFSET
                Case BackslashCharacter
                    writeToBuffer = True: nonInclusiveRegionEnd = characterIndex: newOutputIndex = outputIndex: groupedBackslashCount = 0
                    ' Count how many backslashes there are until a non-backslash character is found.
                    Do
                        groupedBackslashCount = groupedBackslashCount + 1
                        characterIndex = characterIndex + INDEX_OFFSET
                        If characterIndex > Len(json) Then Exit Do
                    Loop Until Mid$(json, characterIndex, 1) <> BackslashCharacter
                    
                    ' Determine how many \ to keep with integer division.
                    countOfBackslashToKeep = groupedBackslashCount \ 2
                    
                    If countOfBackslashToKeep > 0 Then
                        ' Write backslashes to output
                        For newOutputIndex = newOutputIndex + INDEX_OFFSET To newOutputIndex + (2 * countOfBackslashToKeep) Step INDEX_OFFSET
                            Mid$(json, newOutputIndex, 1) = BackslashCharacter
                        Next newOutputIndex
                        ' For loops will exceed boundary on last loop so reverse by the value of step.
                        newOutputIndex = newOutputIndex - INDEX_OFFSET
                    End If
                    ' If there is an even number of backslashes then all of them are escaped and the following character should be ignored.
                    ignoreSpecialCharacter = (groupedBackslashCount Mod 2) = 0
                    
                    If characterIndex <= lastIndex Then

                        currentCharacter = Mid$(json, characterIndex, 1): newOutputIndex = newOutputIndex + INDEX_OFFSET
                        
                        If ignoreSpecialCharacter Then
                            Mid$(json, newOutputIndex, 1) = currentCharacter
                        Else
                            Select Case currentCharacter
                                Case "b" '\b
                                    Mid$(json, newOutputIndex, 1) = VBA.vbBack  'backspace
                                Case "n" '\n
                                    Mid$(json, newOutputIndex, 1) = vbNewLine 'newline
                                Case "t" '\t
                                    Mid$(json, newOutputIndex, 1) = vbTab  'tab
                                Case "0"
                                    'Null character.
                                    newOutputIndex = newOutputIndex - INDEX_OFFSET
                                Case QuoteCharacter, "'", "/" '\" or \/ or \0 or \'
                                    Mid$(json, newOutputIndex, 1) = currentCharacter
                                Case "r" '\r
                                    Mid$(json, newOutputIndex, 1) = vbCr 'carriage return
                                Case "f" '\f
                                    Mid$(json, newOutputIndex, 1) = vbFormFeed ' Form feed
                                Case "u" '\u
                                    ' See: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/chr-function
                                    Mid$(json, newOutputIndex, 1) = ChrW(val("&h" & Mid$(json, characterIndex + 1, 4)))
                                    characterIndex = characterIndex + 4
                                Case Else
                                    Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: Unescaped backslash with no special character."
                            End Select
                        End If
                    End If
                    newRegionStart = characterIndex + INDEX_OFFSET
                Case vbLf, vbCr, vbTab, vbFormFeed
                    writeToBuffer = True: nonInclusiveRegionEnd = characterIndex: newRegionStart = characterIndex + INDEX_OFFSET
                Case Else
                    'For the last index
                    writeToBuffer = True: nonInclusiveRegionEnd = characterIndex + 1
            End Select
            
            If writeToBuffer Or characterIndex = lastIndex Then
                If Not writeToBuffer Then nonInclusiveRegionEnd = characterIndex + INDEX_OFFSET
                outputIndex = outputIndex + 1
                ' [json] is used as both source and buffer.[outputIndex] initializes at 1.
                Mid$(json, outputIndex, nonInclusiveRegionEnd - regionStart) = Mid$(json, regionStart, nonInclusiveRegionEnd - regionStart)
                ' Set to index of last addition.
                outputIndex = outputIndex + (nonInclusiveRegionEnd - regionStart) - 1
                ' If backslash case was used.
                If newOutputIndex > 0 Then outputIndex = newOutputIndex: newOutputIndex = 0
                regionStart = newRegionStart: writeToBuffer = False
            End If
        Loop
        ' Code has been using [json] as a buffer beginning at index 1.
        isRangeQuoteEnclosed = Left$(json, 1) = QuoteCharacter And Mid$(json, outputIndex, 1) = QuoteCharacter
        '
        If Not isRangeQuoteEnclosed Then
            If (Left$(json, 1) = LeftSquare And Mid$(json, outputIndex, 1) = RightSquare) Then
                isDelimiterEnclosed = True
            ElseIf (Left$(json, 1) = LeftCurly And Mid$(json, outputIndex, 1) = RightCurly) Then
                isDelimiterEnclosed = True
            End If
        End If
        ' If there are 2 unescaped quotes in the value region and they occur at the ends, then the segement is a string.
        If (unescapedQuoteCount = 2 And isRangeQuoteEnclosed) Then
            outputString = Mid$(json, 2, outputIndex - 2)
        ElseIf unescapedQuoteCount > 0 And Not isDelimiterEnclosed Then
            ' Ensure quality of string value. ie don't allow "apple to be excepted as a string.
            Err.Raise INVALID_JSON_FORMAT_ERROR, Description:="Invalid Format: String value isn't enclosed in quotes: " & Left$(json, outputIndex)
        Else
            outputString = Left$(json, outputIndex)
        End If
    Else
        isRangeQuoteEnclosed = False
        outputString = Mid$(json, startIndex, lastIndex - startIndex + 1)
    End If

    ProcessCharactersInRange = outputString
    Exit Function
PropagateError:
    PropagateError Err, "ProcessCharactersInRange"
End Function
Private Function TryGetWhitespaceCleanedIndexes(json As String, ByVal regionStartIndex&, ByVal regionEndIndex&, cleanedStartIndex&, cleanedEndIndex&) As Boolean
'========================================================================================================================
'   Summary:
'       Searches for non space,newlines,line feeds and carriage returns from both ends of [json] if they exist.
'       cleanedStartIndex and cleanedEndIndex will be set to the relevant non-white space bytes.
'============================================================================================================
    Dim cleaningLeft As Boolean, cleaningRight As Boolean, errorLeft As Boolean, errorRight As Boolean

    'Max index should be the non null portion of a string byte.
    cleaningLeft = True: cleaningRight = True

    Do
        If cleaningLeft Then
            On Error GoTo CatchLeft
            
            Select Case Mid$(json, regionStartIndex, 1)
                Case SpaceCharacter, vbTab, vbCr, vbLf
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
                Case SpaceCharacter, vbTab, vbCr, vbLf
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

Private Function BinaryCharacterIndexNotWithinQuotes(ByRef json$, ByRef character$, ByRef startIndex&, ByRef maxIndex&) As Long
'===============================================================================================================
' Summary
'   Search for the next [character] not contained within quotes.
' Remarks:
'   [character] within escaped quotes will be ignored.
'===============================================================================================================
    Dim characterIndexBinary&, leftQuoteIndexBinary&, loopCount&, rightQuoteIndexBinary&, _
    adjustLeftQuote As Long, isNotEscapedQuote As Boolean, quoteIndexBinary&, updateCharacterlocation As Boolean, maxIndexBinary&
    
    On Error GoTo Return_Default
    
    maxIndexBinary = (maxIndex * BINARY_OFFSET) - 1
    ' - BINARY_OFFSET so that start position will be maintained when calculating [characterIndexBinary] and following quotes.
    rightQuoteIndexBinary = ((startIndex * BINARY_OFFSET) - 1) - BINARY_OFFSET
    
    updateCharacterlocation = True
    
    Do
        If updateCharacterlocation Then
            characterIndexBinary = InStrB(rightQuoteIndexBinary + BINARY_OFFSET, json, character)
            If characterIndexBinary = 0 Then Exit Do
        End If
        
        loopCount = loopCount + 1
        ' Start searching from the location of right quote: wiil be incremented before searching.
        leftQuoteIndexBinary = rightQuoteIndexBinary

        For adjustLeftQuote = -1 To 0

            quoteIndexBinary = leftQuoteIndexBinary
            Do
                quoteIndexBinary = InStrB(quoteIndexBinary + BINARY_OFFSET, json, QuoteCharacter)
                If Not quoteIndexBinary = 0 Then
                    ' Set to TRUE if previous character in the string isn't a backslash.
                    isNotEscapedQuote = Mid$(json, ((quoteIndexBinary \ BINARY_OFFSET) + 1) - INDEX_OFFSET, 1) <> BackslashCharacter
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
        BinaryCharacterIndexNotWithinQuotes = 0
    ElseIf Not updateCharacterlocation Then
        BinaryCharacterIndexNotWithinQuotes = characterIndexBinary
    End If
    
    Exit Function
Return_Default:
    BinaryCharacterIndexNotWithinQuotes = 0
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
            .source = LeftSquare & .source & RightSquare & delim & parentName
        Else
            sourceParts = Split(.source, delim, 2)
            sourceParts(1) = parentName & " > " & sourceParts(1)
            .source = Join(sourceParts, delim)
        End If
    End With
End Sub
Public Sub Unit_Test()
    Dim success As Boolean ', T As New TimedTask
        Me.ParseJsonString "{""apple"":""tree"",""inner object"":1,""array"":[1,2,3,{""apple"":22,""berry"":45}]}" ', True
        '.ParseJsonString HTTPGet("http://worldtimeapi.org/api/timezone/America/New_York", success) ', True
    'With T.StartSubTask("CFTC")
        'Me.ParseJsonString HTTPGet("https://publicreporting.cftc.gov/resource/jun7-fc8e.json?$select=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name&$group=cftc_contract_market_code,commodity_group_name,commodity_subgroup_name&$limit=5000", success) ', True
        '.EndTask
    'End With
    'T.DPrint
End Sub
