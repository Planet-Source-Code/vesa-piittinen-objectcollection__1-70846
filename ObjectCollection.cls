VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
'* ObjectCollection - Collection compatible replacement for objects
'* ----------------------------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*
'* LICENSE
'* -------
'* http://creativecommons.org/licenses/by-sa/1.0/fi/deed.en
'*
'* Terms: 1) If you make your own version, share using this same license.
'*        2) When used in a program, mention my name in the program's credits.
'*        3) Free for commercial and non-commercial usage.
'*        4) Use at your own risk. No support guaranteed.
'*
'* NOTES
'* -----
'* This class is slower than the native Collection. I mostly wanted to provide a more feature rich
'* collection as well as show how to get the For Each support. To optimize speed the whole logic
'* has to be rebuilt and rethought from scratch.
'*
'* REQUIREMENTS
'* ------------
'* Huh what? Just this one class module. No extra files required.
'*
'* VERSION HISTORY
'* ---------------
'* Version 1.1 (2008-07-18)
'* - Clear method (how many times you've wished to clean a Collection without defining New Collection?)
'* - Clone (easy and fast way to clone the whole collection)
'* - Full Collection syntax compatibility (for faster access use ByIndex, ByKey or ByPtr properties)
'* - Index can be changed
'* - Key can be changed
'* - Key is no longer a requirement
'* - Keys, Items and Ptrs methods to get arrays containing the keys, items or pointers
'* - Swapping
'* - You can change Count (this way you can allocate just what you need to minimize memory usage)
'*
'* Version 1.0 (2008-07-17)
'* - Initial release: supports For Each, Add, Count, Item, Key, Remove
'*
'* CREDITS
'* -------
'* Paul Caton and LaVolpe for their work on SelfSub, SelfHook and SelfCallback
'*************************************************************************************************
Option Explicit

' decrease this value to use less memory: initial usage 98304 bytes (2048 * 4 * 12)
Private Const BUFFERALLOCSIZE = 2048&

Private Type ObjectInfo
    Pos As Long
    KeyLen As Long
    Value As Object
End Type

' public properties
Private m_Count As Long
Private m_Item() As ObjectInfo

' private variables
Private m_Keys As String
Private m_MaxCount As Long

' the following are the requirements for the enumeration
Private Type EnumVariant
    vTable As Long
    RefCount As Long
    Index As Long
    Parent As Object
End Type

Private m_E As EnumVariant
Private m_vTable(6) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Src As Any, ByVal cb As Long)

' the following are the requirements for SelfCallback
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub RtlMachineCodeCopy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Long, Source As MachineCode, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

' self-documentation: less comments and constants
Private Type MachineCode        ' 37 * 4 = 148 bytes
    ' thunk
    OwnerPtr As Long            ' 0
    CallbackAddress As Long     ' 1
    API_EbMode As Long          ' 2
    API_IsBadCodePtr As Long    ' 3
    API_KillTimer As Long       ' 4
    ' code
    MC1(5 To 5) As Long         ' 5
    AllocatedDataPtr As Long    ' 6
    MC2(7 To 17) As Long        ' 7 - 17
    ParamCount As Long          ' 18
    MC3(19 To 35) As Long       ' 19 - 35
    Ordinal As Long             ' 36
End Type

Public Function Add(ByRef Item As Object, Optional ByRef Key As String, Optional ByVal Before As Long, Optional ByVal After As Long) As Long
    Dim lngLen As Long, lngToIndex As Long, EmptyItem As ObjectInfo
    lngLen = Len(Key)
    ' if have key, it must not exist
    If (lngLen <> 0 And InStr(m_Keys, vbNullChar & Key & vbNullChar) = 0) Or (lngLen = 0) Then
        ' determine the zero based index we add the item to
        If Before > 0 And Before <= m_Count Then
            lngToIndex = Before - 1
        ElseIf After > 0 And After < m_Count Then
            lngToIndex = After
        Else
            lngToIndex = m_Count
        End If
        ' allocate more space as necessary
        If m_MaxCount <= m_Count Then
            m_MaxCount = m_MaxCount + BUFFERALLOCSIZE
            ReDim Preserve m_Item(m_MaxCount)
        End If
        ' not setting to last item?
        If lngToIndex < m_Count Then
            ' move old data
            RtlMoveMemory m_Item(lngToIndex + 1), m_Item(lngToIndex), (m_Count - lngToIndex) * 12
            ' zero position before setting
            RtlMoveMemory m_Item(lngToIndex), EmptyItem, 12
        End If
        ' remember key
        If lngLen Then
            m_Item(lngToIndex).KeyLen = lngLen
            m_Item(lngToIndex).Pos = Len(m_Keys) + 1
            ' add to our key string
            m_Keys = m_Keys & (Key & vbNullChar)
        End If
        ' remember the item
        Set m_Item(lngToIndex).Value = Item
        ' increase counter
        m_Count = m_Count + 1
        ' return one based index of the new item
        Add = lngToIndex + 1
    End If
End Function
Public Sub Clear()
    Dim lngA As Long
    ' clean up objects
    For lngA = 0 To m_Count - 1
        Set m_Item(lngA).Value = Nothing
    Next lngA
    ' reset start situatation
    m_Count = 0
    m_MaxCount = BUFFERALLOCSIZE - 1
    ReDim m_Item(m_MaxCount)
    m_Keys = vbNullChar
End Sub
Public Function Clone() As ObjectCollection
    Set Clone = New ObjectCollection
    Clone.Serialize m_Count, m_Item, m_Keys, m_MaxCount
End Function
Public Property Get Count() As Long
    Count = m_Count
End Property
Public Property Let Count(ByVal NewValue As Long)
    Const DoubleNullChar = vbNullChar & vbNullChar
    Dim lngA As Long, lngPos As Long
    If NewValue > 0 Then
        ' we do this the easy way: shrink the buffer out!
        m_Count = NewValue
        m_MaxCount = m_Count - 1
        ReDim Preserve m_Item(m_MaxCount)
        ' get old keys, remove the removed keys
        m_Keys = vbNullChar & Join(Me.Keys, vbNullChar) & vbNullChar
        ' remove empty keys
        Do While InStr(m_Keys, DoubleNullChar)
            m_Keys = Replace(m_Keys, DoubleNullChar, vbNullChar)
        Loop
        ' now reindex the key information
        lngPos = 2
        For lngA = 0 To m_MaxCount
            ' if had a key previously...
            If m_Item(lngA).Pos Then
                ' update the new position (the keys are in the same order now)
                m_Item(lngA).Pos = lngPos
                ' calculate next key position
                lngPos = lngPos + m_Item(lngA).KeyLen + 1
            End If
        Next lngA
    ElseIf NewValue = 0 Then
        ' simple and effective, zero it all out
        Me.Clear
    End If
End Property
Public Property Get Index(ByRef Key As String) As Long
    Dim lngA As Long, lngPos As Long
    ' must have key
    If LenB(Key) Then
        ' find position
        lngPos = InStr(m_Keys, vbNullChar & Key & vbNullChar) + 1
        ' do we have it?
        If lngPos > 1 Then
            ' now we just have to find which item it is
            For lngA = 0 To m_Count - 1
                ' it is more likely to not have a match: True condition is faster
                If m_Item(lngA).Pos <> lngPos Then
                    ' not found
                Else
                    ' found!
                    Index = lngA + 1
                    Exit Property
                End If
            Next lngA
        End If
    End If
End Property
Public Property Let Index(ByRef Key As String, ByVal NewValue As Long)
    Dim lngA As Long, lngPos As Long, TempItem As ObjectInfo, EmptyItem As ObjectInfo
    ' validate index & must have key
    If NewValue >= 1 And NewValue <= m_Count And LenB(Key) <> 0 Then
        ' convert to zero base
        NewValue = NewValue - 1
        ' find position
        lngPos = InStr(m_Keys, vbNullChar & Key & vbNullChar) + 1
        ' do we have it?
        If lngPos > 1 Then
            ' now we just have to find which item it is
            For lngA = 0 To m_Count - 1
                ' it is more likely to not have a match: True condition is faster
                If m_Item(lngA).Pos <> lngPos Then
                    ' not found
                Else
                    ' found!
                    If lngA < NewValue Then
                        ' move to a later position
                        RtlMoveMemory TempItem, m_Item(lngA), 12
                        RtlMoveMemory m_Item(lngA), m_Item(lngA + 1), (NewValue - lngA) * 12
                        RtlMoveMemory m_Item(NewValue), TempItem, 12
                        RtlMoveMemory TempItem, EmptyItem, 12
                    ElseIf lngA > NewValue Then
                        ' move to an earlier position
                        RtlMoveMemory TempItem, m_Item(lngA), 12
                        RtlMoveMemory m_Item(NewValue + 1), m_Item(NewValue), (lngA - NewValue) * 12
                        RtlMoveMemory m_Item(NewValue), TempItem, 12
                        RtlMoveMemory TempItem, EmptyItem, 12
                    End If
                    Exit Property
                End If
            Next lngA
        End If
    End If
End Property
Public Property Get IndexByPtr(ByVal Ptr As Long) As Long
    Dim lngA As Long
    ' loop through all items
    For lngA = 0 To m_Count - 1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            IndexByPtr = lngA + 1
        End If
    Next lngA
End Property
Public Property Let IndexByPtr(ByVal Ptr As Long, ByVal NewValue As Long)
    Dim lngA As Long, TempItem As ObjectInfo, EmptyItem As ObjectInfo
    ' validate index & must have key
    If NewValue >= 1 And NewValue <= m_Count Then
        ' convert to zero base
        NewValue = NewValue - 1
        ' loop through all items
        For lngA = 0 To m_Count - 1
            ' see if pointer matches
            If ObjPtr(m_Item(lngA).Value) <> Ptr Then
                ' not found
            Else
                ' found!
                If lngA < NewValue Then
                    ' move to a later position
                    RtlMoveMemory TempItem, m_Item(lngA), 12
                    RtlMoveMemory m_Item(lngA), m_Item(lngA + 1), (NewValue - lngA) * 12
                    RtlMoveMemory m_Item(NewValue), TempItem, 12
                    RtlMoveMemory TempItem, EmptyItem, 12
                ElseIf lngA > NewValue Then
                    ' move to an earlier position
                    RtlMoveMemory TempItem, m_Item(lngA), 12
                    RtlMoveMemory m_Item(NewValue + 1), m_Item(NewValue), (lngA - NewValue) * 12
                    RtlMoveMemory m_Item(NewValue), TempItem, 12
                    RtlMoveMemory TempItem, EmptyItem, 12
                End If
                Exit Property
            End If
        Next lngA
    End If
End Property
Public Property Get Item(Index) As Object
Attribute Item.VB_UserMemId = 0
    Dim lngType As Long
    ' get the variant type
    lngType = VarType(Index)
    ' make sure not an array
    If (lngType And vbArray) = 0 Then
        Select Case lngType
            ' string is straightforward
            Case vbString
                Set Item = ItemByKey(CStr(Index))
            ' allow any numeric datatype
            Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                Set Item = ItemByIndex(CLng(Index))
        End Select
    End If
End Property
Public Property Set Item(Index, ByRef NewValue As Object)
    Dim lngType As Long
    ' get the variant type
    lngType = VarType(Index)
    ' make sure not an array
    If (lngType And vbArray) = 0 Then
        Select Case lngType
            ' string is straightforward
            Case vbString
                Set ItemByKey(CStr(Index)) = NewValue
            ' allow any numeric datatype
            Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                Set ItemByIndex(Index) = NewValue
        End Select
    End If
End Property
Public Property Get ItemByIndex(ByVal Index As Long) As Object
    If Index >= 1 And Index <= m_Count Then
        Set ItemByIndex = m_Item(Index - 1).Value
    End If
End Property
Public Property Set ItemByIndex(ByVal Index As Long, ByRef NewValue As Object)
    If Index >= 1 And Index <= m_Count Then
        Set m_Item(Index - 1).Value = NewValue
    End If
End Property
Public Property Get ItemByPtr(ByVal Ptr As Long) As Object
    Dim lngA As Long
    ' loop through all items
    For lngA = 0 To m_Count - 1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            Set ItemByPtr = m_Item(lngA).Value
        End If
    Next lngA
End Property
Public Property Set ItemByPtr(ByVal Ptr As Long, ByRef NewValue As Object)
    Dim lngA As Long
    ' loop through all items
    For lngA = 0 To m_Count - 1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            Set ItemByIndex(lngA + 1) = NewValue
            Exit Property
        End If
    Next lngA
End Property
Public Property Get ItemByKey(ByRef Key As String) As Object
    Dim lngA As Long, lngPos As Long
    ' must have key
    If LenB(Key) Then
        ' find position
        lngPos = InStr(m_Keys, vbNullChar & Key & vbNullChar) + 1
        ' do we have it?
        If lngPos > 1 Then
            ' now we just have to find which item it is
            For lngA = 0 To m_Count - 1
                With m_Item(lngA)
                    ' it is more likely to not have a match: True condition is faster
                    If .Pos <> lngPos Then
                        ' not found
                    Else
                        ' found!
                        Set ItemByKey = .Value
                        Exit Property
                    End If
                End With
            Next lngA
        End If
    End If
End Property
Public Property Set ItemByKey(ByRef Key As String, ByRef NewValue As Object)
    Dim lngA As Long, lngPos As Long
    ' must have key
    If LenB(Key) Then
        ' find position
        lngPos = InStr(m_Keys, vbNullChar & Key & vbNullChar) + 1
        ' do we have it?
        If lngPos > 1 Then
            ' now we just have to find which item it is
            For lngA = 0 To m_Count - 1
                With m_Item(lngA)
                    ' it is more likely to not have a match: True condition is faster
                    If .Pos <> lngPos Then
                        ' not found
                    Else
                        ' found!
                        Set .Value = NewValue
                        Exit Property
                    End If
                End With
            Next lngA
        End If
    End If
End Property
Public Function Items(Optional ByVal Base As Long = 0) As Object()
    Dim lngA As Long, objOut() As Object
    ' must have items
    If m_Count Then
        ' create the output array using the given base
        ReDim objOut(Base To m_Count - 1 + Base)
        ' fill the array with new references to the objects
        For lngA = Base To UBound(objOut)
            Set objOut(lngA) = m_Item(lngA - Base).Value
        Next lngA
        ' return the array
        Items = objOut
    End If
End Function
Public Property Get Key(ByVal Index As Long) As String
    Dim lngA As Long, lngPos As Long
    If Index >= 1 And Index <= m_Count Then
        ' get the key for this index
        With m_Item(Index - 1)
            If .Pos Then Key = Mid$(m_Keys, .Pos, .KeyLen)
        End With
    End If
End Property
Public Property Let Key(ByVal Index As Long, ByRef NewValue As String)
    Dim lngA As Long, lngLen As Long, lngPos As Long
    If Index >= 1 And Index <= m_Count Then
        ' validate the key does not exist already
        If InStr(m_Keys, vbNullChar & NewValue & vbNullChar) = 0 Then
            ' lock the array
            With m_Item(Index - 1)
                ' the existing item has a key?
                If .Pos Then
                    ' replace or remove?
                    If LenB(NewValue) Then
                        ' replace the key
                        m_Keys = Left$(m_Keys, .Pos - 1) & NewValue & Mid$(m_Keys, .Pos + .KeyLen)
                        ' cache the position to compare to
                        lngPos = .Pos
                        ' get the length difference
                        lngLen = .KeyLen - Len(NewValue)
                    Else
                        ' remove the key
                        m_Keys = Left$(m_Keys, .Pos - 1) & Mid$(m_Keys, .Pos + .KeyLen + 1)
                        ' cache the position to compare to
                        lngPos = .Pos
                        ' removed key length
                        lngLen = .KeyLen + 1
                        ' remove position information
                        .Pos = 0
                    End If
                    ' update key length
                    .KeyLen = Len(NewValue)
                    ' if key length differs...
                    If lngLen Then
                        ' we has a loopy situatations
                        For lngA = 0 To m_Count - 1
                            ' if a later position...
                            If m_Item(lngA).Pos > lngPos Then
                                ' update the position information
                                m_Item(lngA).Pos = m_Item(lngA).Pos - lngLen
                            End If
                        Next lngA
                    End If
                    Exit Property
                ElseIf LenB(NewValue) Then
                    ' add a new key
                    .Pos = Len(m_Keys) + 1
                    .KeyLen = Len(NewValue)
                    m_Keys = m_Keys & (NewValue & vbNullChar)
                    Exit Property
                Else
                    ' no key to no key, no change
                    Exit Property
                End If
            End With
        End If
    End If
End Property
Public Property Get KeyByPtr(ByVal Ptr As Long) As String
    Dim lngA As Long
    ' loop through all items
    For lngA = 0 To m_Count - 1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            If m_Item(lngA).Pos Then KeyByPtr = Mid$(m_Keys, m_Item(lngA).Pos, m_Item(lngA).KeyLen)
        End If
    Next lngA
End Property
Public Property Let KeyByPtr(ByVal Ptr As Long, ByRef NewValue As String)
    Dim lngA As Long
    ' loop through all items
    For lngA = 0 To m_Count - 1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            Me.Key(lngA + 1) = NewValue
            Exit Property
        End If
    Next lngA
End Property
Public Function Keys(Optional ByVal Base As Long = 0) As String()
    Dim lngA As Long, strOut() As String
    ' must have items
    If m_Count Then
        ' create the output array using the given base
        ReDim strOut(Base To m_Count - 1 + Base)
        ' fill the array with the keys in order
        For lngA = Base To UBound(strOut)
            With m_Item(lngA - Base)
                ' set key to array
                If .Pos Then strOut(lngA) = Mid$(m_Keys, .Pos, .KeyLen)
            End With
        Next lngA
        ' return the array
        Keys = strOut
    End If
End Function
Public Function Ptrs(Optional ByVal Base As Long = 0) As Long()
    Dim lngA As Long, lngOut() As Long
    ' must have items
    If m_Count Then
        ' create the output array using the given base
        ReDim lngOut(Base To m_Count - 1 + Base)
        ' fill the array with the pointers in order
        For lngA = Base To UBound(lngOut)
            ' set pointer to array
            lngOut(lngA) = ObjPtr(m_Item(lngA - Base).Value)
        Next lngA
        ' return the array
        Ptrs = lngOut
    End If
End Function
Public Function Remove(Index) As Boolean
    Dim lngType As Long
    ' get the variant type
    lngType = VarType(Index)
    ' make sure not an array
    If (lngType And vbArray) = 0 Then
        Select Case lngType
            ' string is straightforward
            Case vbString
                Remove = RemoveByKey(CStr(Index))
            ' allow any numeric datatype
            Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                Remove = RemoveByIndex(CLng(Index))
        End Select
    End If
End Function
Public Function RemoveByIndex(ByVal Index As Long) As Boolean
    Dim lngA As Long, lngB As Long, lngLen As Long, lngPos As Long, EmptyItem As ObjectInfo
    If Index >= 1 And Index <= m_Count Then
        lngA = Index - 1
        Set m_Item(lngA).Value = Nothing
        ' key length
        lngLen = m_Item(lngA).KeyLen + 1
        ' key position
        lngPos = m_Item(lngA).Pos
        ' decrease count
        m_Count = m_Count - 1
        If lngA < m_Count Then
            ' move items
            RtlMoveMemory m_Item(lngA), m_Item(lngA + 1), (m_Count - lngA) * 12
            RtlMoveMemory m_Item(m_Count), EmptyItem, 12
        End If
        ' correct indexes
        For lngB = 0 To m_Count - 1
            If m_Item(lngB).Pos > lngPos Then m_Item(lngB).Pos = m_Item(lngB).Pos - lngLen
        Next lngB
        ' remove the key
        m_Keys = Left$(m_Keys, lngPos - 1) & Mid$(m_Keys, lngPos + lngLen)
        ' success!
        RemoveByIndex = True
        Exit Function
    End If
End Function
Public Function RemoveByKey(ByRef Key As String) As Boolean
    Dim lngA As Long, lngB As Long, lngLen As Long, lngPos As Long, EmptyItem As ObjectInfo
    ' must have key
    If LenB(Key) Then
        ' find position
        lngPos = InStr(m_Keys, vbNullChar & Key & vbNullChar) + 1
        ' do we have it?
        If lngPos > 1 Then
            ' now we just have to find which item it is
            For lngA = 0 To m_Count - 1
                ' it is more likely to not have a match: True condition is faster
                If m_Item(lngA).Pos <> lngPos Then
                    ' not found
                Else
                    ' found!
                    Set m_Item(lngA).Value = Nothing
                    ' key length
                    lngLen = m_Item(lngA).KeyLen + 1
                    ' decrease count
                    m_Count = m_Count - 1
                    If lngA < m_Count Then
                        ' move items
                        RtlMoveMemory m_Item(lngA), m_Item(lngA + 1), (m_Count - lngA) * 12
                        RtlMoveMemory m_Item(m_Count), EmptyItem, 12
                    End If
                    ' correct indexes
                    For lngB = 0 To m_Count - 1
                        If m_Item(lngB).Pos > lngPos Then m_Item(lngB).Pos = m_Item(lngB).Pos - lngLen
                    Next lngB
                    ' remove the key
                    m_Keys = Left$(m_Keys, lngPos - 1) & Mid$(m_Keys, lngPos + lngLen)
                    ' success!
                    RemoveByKey = True
                    Exit Function
                End If
            Next lngA
        End If
    End If
End Function
Public Function RemoveByPtr(ByVal Ptr As Long) As Boolean
    Dim lngA As Long
    ' loop through all items (note: all copies of the same object are removed)
    For lngA = m_Count - 1 To 0 Step -1
        ' see if pointer matches
        If ObjPtr(m_Item(lngA).Value) <> Ptr Then
            ' not found
        Else
            ' found!
            RemoveByPtr = Me.RemoveByIndex(lngA + 1)
        End If
    Next lngA
End Function
Friend Sub Serialize(ByVal Count As Long, ByRef Items() As ObjectInfo, ByRef Keys As String, ByVal MaxCount As Long)
    m_Count = Count
    m_Item = Items
    m_Keys = Keys
    m_MaxCount = MaxCount
End Sub
Public Function Swap(Index1, Index2) As Boolean
    Dim lngA As Long, lngPos As Long, lngType As Long
    ' get the variant type
    lngType = VarType(Index1)
    ' make sure not an array
    If (lngType And vbArray) = 0 Then
        Select Case lngType
            Case vbString
                If LenB(Index1) Then
                    lngPos = InStr(m_Keys, vbNullChar & Index1 & vbNullChar) + 1
                    If lngPos > 1 Then
                        ' now we just have to find which item it is
                        For lngA = 0 To m_Count - 1
                            ' it is more likely to not have a match: True condition is faster
                            If m_Item(lngA).Pos <> lngPos Then
                                ' not found
                            Else
                                ' found!
                                Index1 = lngA + 1
                                Exit For
                            End If
                        Next lngA
                        If m_Count = lngA Then Exit Function
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            ' allow any numeric datatype
            Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                ' force Long
                Index1 = CLng(Index1)
            Case Else
                Exit Function
        End Select
    Else
        Exit Function
    End If
    ' get the variant type
    lngType = VarType(Index2)
    ' make sure not an array
    If (lngType And vbArray) = 0 Then
        Select Case lngType
            Case vbString
                If LenB(Index2) Then
                    lngPos = InStr(m_Keys, vbNullChar & Index2 & vbNullChar) + 1
                    If lngPos > 1 Then
                        ' now we just have to find which item it is
                        For lngA = 0 To m_Count - 1
                            ' it is more likely to not have a match: True condition is faster
                            If m_Item(lngA).Pos <> lngPos Then
                                ' not found
                            Else
                                ' found!
                                Index2 = lngA + 1
                                Exit For
                            End If
                        Next lngA
                        If m_Count = lngA Then Exit Function
                    Else
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            ' allow any numeric datatype
            Case vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
                ' force Long
                Index2 = CLng(Index2)
            Case Else
                Exit Function
        End Select
    Else
        Exit Function
    End If
    ' now we finally have indexes!
    Swap = SwapByIndex(Index1, Index2)
End Function
Public Function SwapByIndex(ByVal Index1 As Long, ByVal Index2 As Long) As Boolean
    Dim lngLen As Long, lngPos As Long
    ' validate indexes
    If (Index1 >= 1 And Index1 <= m_Count) And (Index2 >= 1 And Index2 <= m_Count) And (Index1 <> Index2) Then
        ' convert to zero base
        Index1 = Index1 - 1
        Index2 = Index2 - 1
        ' swap key information
        lngLen = m_Item(Index1).KeyLen
        lngPos = m_Item(Index1).Pos
        m_Item(Index1).KeyLen = m_Item(Index2).KeyLen
        m_Item(Index1).Pos = m_Item(Index2).Pos
        m_Item(Index2).KeyLen = lngLen
        m_Item(Index2).Pos = lngPos
        ' success!
        SwapByIndex = True
        Exit Function
    ElseIf Index1 = Index2 Then
        ' same index, return true
        SwapByIndex = True
        Exit Function
    End If
End Function
Public Function SwapByKey(ByRef Key1 As String, ByRef Key2 As String) As Boolean
    Dim lngA As Long, lngLen As Long, lngPos1 As Long, lngPos2 As Long, Index1 As Long, Index2 As Long
    If (LenB(Key1) <> 0) And (LenB(Key2) <> 0) Then
        ' get indexes
        lngPos1 = InStr(m_Keys, vbNullChar & Key1 & vbNullChar) + 1
        lngPos2 = InStr(m_Keys, vbNullChar & Key2 & vbNullChar) + 1
        ' if both were found...
        If (lngPos1 > 1) And (lngPos2 > 1) Then
            ' if same key we return true and exit
            If lngPos1 = lngPos2 Then SwapByKey = True: Exit Function
            ' invalid indexes
            Index1 = m_Count
            Index2 = m_Count
            ' we has a loopy situatations
            For lngA = 0 To m_Count - 1
                If lngPos1 = m_Item(lngA).Pos Then
                    Index1 = lngA
                ElseIf lngPos2 = m_Item(lngA).Pos Then
                    Index2 = lngA
                End If
                ' True condition is faster and is more likely
                If (Index1 = m_Count) Or (Index2 = m_Count) Then Else Exit For
            Next lngA
            ' if we found both indexes...
            If lngA < m_Count Then
                ' we can swap
                lngLen = m_Item(Index1).KeyLen
                m_Item(Index1).KeyLen = m_Item(Index2).KeyLen
                m_Item(Index1).Pos = lngPos2
                m_Item(Index2).KeyLen = lngLen
                m_Item(Index2).Pos = lngPos1
                ' success!
                SwapByKey = True
                Exit Function
            End If
        End If
    End If
End Function
Public Function SwapByPtr(ByVal Ptr1 As Long, ByVal Ptr2 As Long) As Boolean
    Dim lngA As Long, lngPtr As Long, Index1 As Long, Index2 As Long
    If Ptr1 <> Ptr2 Then
        ' loop through all items
        For lngA = 0 To m_Count - 1
            ' see if pointer matches
            lngPtr = ObjPtr(m_Item(lngA).Value)
            If lngPtr = Ptr1 Then
                Index1 = lngA + 1
            ElseIf lngPtr = Ptr2 Then
                Index2 = lngA + 1
            End If
            ' continue while either one is invalid index
            If (Index1 = 0) Or (Index2 = 0) Then Else Exit For
        Next lngA
        ' got the indexes?
        If lngA < m_Count Then SwapByPtr = Me.SwapByIndex(Index1, Index2)
    Else
        SwapByPtr = True
    End If
End Function
' Tools > Procedure Attributes > NewEnum: Procedure ID = -4 and Hide this member = True
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' return the EnumVariant structure
    RtlMoveMemory NewEnum, VarPtr(m_E), 4
End Function
Private Function Private_AddressOfOrdinal(ByVal Ordinal As Long, ByVal ParamCount As Long) As Long
    Dim bytValue As Byte, bytSignature As Byte, lngA As Long
    Dim lngAddress As Long, lngMethod As Long, lngObject As Long
    Dim lngCallback As Long, lngMCmem As Long, udtMC As MachineCode

    ' get object pointer
    lngObject = ObjPtr(Me)
    ' get VTable address
    GetMem4 lngObject, lngAddress
    ' Class method (see SelfSub code for non-Class values)
    lngAddress = lngAddress + &H1C&
    ' get method pointer
    GetMem4 lngAddress, lngMethod
    ' get method signature byte: &H33 = pseudo-code, &HE9 = native code
    GetMem1 lngMethod, bytSignature
    ' next VTable address
    lngAddress = lngAddress + 4&
    ' try a "reasonable" amount of VTable entries
    For lngA = 511 To 1 Step -1
        ' get method pointer
        GetMem4 lngAddress, lngMethod
        ' see if we are out of VTable (I use "Then Else" because True conditions are faster)
        If IsBadCodePtr(lngMethod) = 0& Then Else Exit For
        ' get method signature byte
        GetMem1 lngMethod, bytValue
        ' if it is invalid we are out of VTable
        If bytValue = bytSignature Then Else Exit For
        ' try next one
        lngAddress = lngAddress + 4&
    Next lngA
    ' if lngA = 0 we looped through the entire loop; if that did not happen, we get the pointer
    If lngA Then GetMem4 lngAddress - (Ordinal * 4&), lngCallback
    ' verify we got the callback address
    If lngCallback Then
        ' allocate executable memory
        lngMCmem = VirtualAlloc(0, LenB(udtMC), &H1000&, &H40&) 'Length, MEM_COMMIT, PAGE_RWX
        ' verify we got it
        If lngMCmem Then
            With udtMC
                ' thunk
                .OwnerPtr = lngObject
                .CallbackAddress = lngCallback
                If App.LogMode = 0 Then
                    ' for IDE safety, store the EbMode function address in the thunk data
                    .API_EbMode = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")
                End If
                .API_IsBadCodePtr = GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
                '.API_KillTimer = GetProcAddress(GetModuleHandleA("user32"), "KillTimer")
                ' actual machine code
                .MC1(5&) = &HBB60E089:    .MC2(7&) = &H73FFC589
                .MC2(8&) = &HC53FF04:     .MC2(9&) = &H59E80A74
                .MC2(10) = &HE9000000:    .MC2(11) = &H30&
                .MC2(12) = &H87B81:       .MC2(13) = &H75000000
                .MC2(14) = &H9090902B:    .MC2(15) = &H42DE889
                .MC2(16) = &H50000000:    .MC2(17) = &HB9909090
                .MC3(19) = &H90900AE3:    .MC3(20) = &H8D74FF
                .MC3(21) = &H9090FAE2:    .MC3(22) = &H53FF33FF
                .MC3(23) = &H90909004:    .MC3(24) = &H2BADC261
                .MC3(25) = &H3D0853FF:    .MC3(26) = &H1&
                .MC3(27) = &H23DCE74:     .MC3(28) = &H74000000
                .MC3(29) = &HAE807:       .MC3(30) = &H90900000
                .MC3(31) = &H4589C031:    .MC3(32) = &H90DDEBFC
                .MC3(33) = &HFF0C75FF:    .MC3(34) = &H53FF0475
                .MC3(35) = &HC310&
                ' settings within the code
                .AllocatedDataPtr = lngMCmem
                .Ordinal = Ordinal
                .ParamCount = ParamCount
                PutMem2 VarPtr(.MC3(24)) + 2&, CInt(ParamCount * 4&)
            End With
            ' copy thunk code to executable memory
            RtlMachineCodeCopy ByVal lngMCmem, udtMC, LenB(udtMC)
            ' return the procedure address (add thunk offset)
            Private_AddressOfOrdinal = lngMCmem + &H14&
        End If
    End If
End Function
Private Sub Private_Start()
    ' initialize
    m_E.vTable = VarPtr(m_vTable(0))
    m_E.RefCount = 1
    m_E.Index = 0
    Set m_E.Parent = Me
    ' fill the vTable
    If m_vTable(0) = 0 Then
        m_vTable(0) = Private_AddressOfOrdinal(7, 3)
        m_vTable(1) = Private_AddressOfOrdinal(6, 1)
        m_vTable(2) = Private_AddressOfOrdinal(5, 1)
        m_vTable(3) = Private_AddressOfOrdinal(4, 4)
        m_vTable(4) = 0 'Private_AddressOfOrdinal(3, 2)
        m_vTable(5) = 0 'Private_AddressOfOrdinal(2, 1)
        m_vTable(6) = 0 'Private_AddressOfOrdinal(1, 1)
    End If
End Sub
Private Sub Private_Stop()
    Dim lngA As Long
    ' free the procedure callbacks
    For lngA = 0 To UBound(m_vTable)
        If m_vTable(lngA) Then VirtualFree m_vTable(lngA), 0&, &H8000& 'MEM_RELEASE
    Next lngA
    ' nullify
    Erase m_vTable
    ' nullify
    m_E.vTable = 0
    m_E.Index = 0
    m_E.RefCount = 0
    ' remove the parent reference
    If Not m_E.Parent Is Nothing Then Set m_E.Parent = Nothing
End Sub
Private Sub Class_Initialize()
    ' allow us to use Not Not Array to get an array pointer without errors in the IDE
    ' special note: you can even get the correct string array pointer!
    Dim IDEbug() As Long
    Debug.Assert Not IDEbug Or App.hInstance
    ' reserve items
    m_MaxCount = BUFFERALLOCSIZE - 1
    ReDim m_Item(m_MaxCount)
    ' init key string
    m_Keys = vbNullChar
    Private_Start
End Sub
Private Sub Class_Terminate()
    Dim lngA As Long
    Private_Stop
    ' clean up objects
    For lngA = 0 To m_Count - 1
        Set m_Item(lngA).Value = Nothing
    Next lngA
    ' clean up everything else
    Erase m_Item
    m_Keys = vbNullString
End Sub
' ordinal #7!
Private Function Variant_QueryInterface(ByRef E As EnumVariant, ByRef riid As Long, ByRef lObj As Long) As Long
    E.RefCount = E.RefCount + 1
    lObj = VarPtr(E)
End Function
' ordinal #6!
Private Function Variant_AddRef(ByRef E As EnumVariant) As Long
    E.RefCount = E.RefCount + 1
    Variant_AddRef = E.RefCount
End Function
' ordinal #5!
Private Function Variant_Release(ByRef E As EnumVariant) As Long
    E.RefCount = E.RefCount - 1
    Variant_Release = E.RefCount
    If E.RefCount > 0 Then Else Set E.Parent = Nothing
End Function
' ordinal #4!
Private Function Variant_Next(ByRef E As EnumVariant, ByVal C As Long, ByRef V As Variant, ByVal pCFtch As Long) As Long
    Dim lngIndex As Long
    lngIndex = E.Index
    pCFtch = C
    If (Not E.Parent Is Nothing) And (m_Count > E.Index) Then
        Set V = m_Item(E.Index).Value: E.Index = E.Index + 1
    End If
    Variant_Next = Abs((E.Index - lngIndex) < C)
End Function
' ordinal #3! UNUSED (correct functionality not guaranteed)
Private Function Variant_Skip(ByRef E As EnumVariant, ByVal Skip As Long) As Long
    If E.Index + Skip < m_Count Then
        E.Index = E.Index + Skip
    Else
        Variant_Skip = 1
    End If
End Function
' ordinal #2! UNUSED (correct functionality not guaranteed)
Private Function Variant_Reset(ByRef E As EnumVariant) As Long
    E.Index = 0
End Function
' ordinal #1! UNUSED (correct functionality not guaranteed)
Private Function Variant_Clone(ByRef E As EnumVariant) As EnumVariant
    Variant_Clone = E
End Function
