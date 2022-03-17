VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a list of elements that can be accessed by index. Provides methods to search, sort, and manipulate lists. \r\n\r\nAuthor: Andrew Tokarskiy tokarskiy.a@keemail.me \r\nGithub: https://github.com/tokarskiy \r\nLicense: http://www.opensource.org/licenses/mit-license.php"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' List v0.4.0
'''
''' Represents a list of elements that can be accessed by index.
''' Provides methods to search, sort and manipulate lists.
'''
''' Author : Andrew Tokarskiy <tokarskiy.a@keemail.me>
''' Github : https://github.com/tokarskiy
''' License: http://www.opensource.org/licenses/mit-license.php
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

#Const UseExcelObjects = False
#Const UseDictionaryWin = False
#Const UseDictionaryMac = False

#Const Win = Win16 Or Win32 Or Win64
#Const UseDictionary = (UseDictionaryMac And Mac) Or (UseDictionaryWin And Win)

Private Const MinIndex As Long = 1
Private seq_pCollection As VBA.Collection

'''
''' Returns the number of items in the List
'''
''' @return {Long}
'''
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the List"
    Count = seq_pCollection.Count
End Property

'''
''' Returns the number that represents how many elements equal to the function parameter
'''
''' @param {Variant} Item
''' @return {Long}
'''
Public Function CountOf(Item As Variant) As Long
Attribute CountOf.VB_Description = "Returns the number that represents how many elements equal to the function parameter"
    Dim elem As Variant
    
    CountOf = 0
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Item) Then
            CountOf = CountOf + 1
        End If
    Next elem
End Function

'''
''' Returns an item by index
'''
''' @param {Long} Index
''' @return {Variant}
'''
Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Dim collIndex As Long
    collIndex = GetCollectionIndex(Index)
    
    If VBA.IsObject(seq_pCollection.Item(collIndex)) Then
        Set Item = seq_pCollection.Item(collIndex)
    Else
        Item = seq_pCollection.Item(collIndex)
    End If
End Property

'''
''' Bypasses a specified number of elements in a List and
''' then returns the remaining elements.
'''
''' @param {Long} Amount
''' @return {List}
'''
Public Function Skip(ByVal Amount As Long) As List
Attribute Skip.VB_Description = "Bypasses a specified number of elements in a List and then returns the remaining elements."
    Dim elem As Variant
    Dim i As Long
    Set Skip = New List
    
    i = 1
    For Each elem In seq_pCollection
        If i > Amount Then
            Skip.Add elem
        End If
        
        i = i + 1
    Next elem
End Function

'''
''' Returns a specified number of contiguous elements
''' from the start of a List.
'''
''' @param {Long} Amount
''' @return {List}
'''
Public Function Take(ByVal Amount As Long) As List
Attribute Take.VB_Description = "Returns a specified number of contiguous elements from the start of a List."
    Dim elem As Variant
    Dim i As Long
    Set Take = New List
    
    i = 0
    For Each elem In seq_pCollection
        If i < Amount Then
            i = i + 1
            Take.Add elem
        End If
    Next elem
End Function

'''
''' Retrieves a sublist from this List instance. The sublist starts at a
''' specified element position and has a specified length.
'''
''' @param {Long} Start
''' @param {Long} Length
''' @return {List}
'''
Public Function Mid(ByVal Start As Long, ByVal Length As Long) As List
Attribute Mid.VB_Description = "Retrieves a sublist from this List instance. The sublist starts at a specified element position and has a specified length."
    Dim elem As Variant
    Dim i As Long
    Dim startCollIndex As Long
    Dim endCollIndex As Long
    
    Set Mid = New List
    startCollIndex = GetCollectionIndex(Start)
    endCollIndex = VBA.IIf(startCollIndex + Length - 1 > seq_pCollection.Count, _
                           seq_pCollection.Count, _
                           startCollIndex + Length - 1)
                               
    If startCollIndex < 1 Or startCollIndex > seq_pCollection.Count Then
        Err.Raise Number:=9 'Subscript out of range
        Exit Function
    End If
    
    For i = startCollIndex To endCollIndex Step 1
        Mid.Add seq_pCollection.Item(i)
    Next i
End Function

'''
''' Returns a specified number of contiguous elements
''' from the start of a List.
'''
''' @param {Long} Amount
''' @return {List}
'''
Public Function Left(ByVal Amount As Long) As List
Attribute Left.VB_Description = "Returns a specified number of contiguous elements from the start of a List"
    Set Left = Take(Amount)
End Function

'''
''' Returns a specified number of contiguous elements
''' from the end of a List.
'''
''' @param {Long} Amount
''' @return {List}
'''
Public Function Right(ByVal Amount As Long) As List
Attribute Right.VB_Description = "Returns a specified number of contiguous elements from the end of a List"
    Set Right = Me.Copy()
    
    Do While Right.Count > Amount
        Right.Remove MinIndex
    Loop
End Function

'''
''' Inserts an element to the List
'''
''' @param {Variant} Item
''' @param {Long} Before
''' @param {Long} After
'''
Public Sub Add(Item As Variant, Optional ByVal Before As Long = MinIndex - 1, Optional ByVal After As Long = MinIndex - 1)
Attribute Add.VB_Description = "Inserts an element to the List"
    Dim collIndex As Long
    If After >= MinIndex Then
        collIndex = GetCollectionIndex(After)
        seq_pCollection.Add Item, After:=collIndex
    ElseIf Before >= MinIndex Then
        collIndex = GetCollectionIndex(Before)
        seq_pCollection.Add Item, Before:=collIndex
    Else
        seq_pCollection.Add Item
    End If
End Sub

'''
''' Appends the elements of specified List to this List
'''
''' @param {List} Sequence
'''
Public Sub AddList(ByRef Sequence As List)
Attribute AddList.VB_Description = "Appends the elements of specified List to this List"
    Dim elem As Variant
    Dim i As Long
    Dim cnt As Long
    
    If Sequence Is Nothing Then
        Exit Sub
    End If
    
    If Sequence Is Me Then
        cnt = Count
        For i = 1 To cnt Step 1
            seq_pCollection.Add Item(i)
        Next i
        Exit Sub
    End If
    
    For Each elem In Sequence
        seq_pCollection.Add elem
    Next elem
End Sub

'''
''' Appends the elements of specified VBA.Collection to this List
'''
''' @param {VBA.Collection} Sequence
'''
Public Sub AddCollection(ByRef Sequence As VBA.Collection)
Attribute AddCollection.VB_Description = "Appends the elements of specified Collection to this List"
    Dim elem As Variant
    Dim i As Long
    Dim cnt As Long
    
    If Sequence Is Nothing Then
        Exit Sub
    End If
    
    For Each elem In Sequence
        seq_pCollection.Add elem
    Next elem
End Sub

'''
''' Appends the elements of specified Excel Range to this List
'''
''' @param {Excel.Range} Range
'''
#If UseExcelObjects Then
Public Sub AddRange(ByRef Range As Excel.Range)
Attribute AddRange.VB_Description = "Appends the elements of specified Excel Range to this List"
    Dim c As Object
    
    For Each c In Range
        seq_pCollection.Add c.Value
    Next c
    
    Set c = Nothing
End Sub
#End If

'''
''' Appends the elements of specified array to this List
'''
''' @param {Variant} Elements
'''
Public Sub AddArray(Elements As Variant)
Attribute AddArray.VB_Description = "Appends the elements of specified array to this List"
    Dim i As Long
    If Not VBA.IsArray(Elements) Then
        VBA.Err.Raise Number:=13 ' Type mismatch
        Exit Sub
    End If

    For i = LBound(Elements) To UBound(Elements) Step 1
        seq_pCollection.Add Elements(i)
    Next i
End Sub

'''
''' Determines whether any element of a List equals the Parameter
'''
''' @param {Variant} Item
''' @return {Boolean}
'''
Public Function Contains(Item As Variant) As Boolean
Attribute Contains.VB_Description = "Determines whether any element of a List equals the Parameter"
    Contains = IndexOf(Item) >= MinIndex
End Function

'''
''' Returns the index of a first element that equals to Parameter
'''
''' @param {Variant} Item
''' @return {Long}
'''
Public Function IndexOf(Item As Variant) As Long
Attribute IndexOf.VB_Description = "Returns the index of a first element that equals to Parameter"
    Dim i As Long
    Dim elem As Variant
    
    i = MinIndex
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Item) Then
            IndexOf = i
            Exit Function
        End If
        
        i = i + 1
    Next elem
    
    IndexOf = MinIndex - 1
End Function

'''
''' Concatenates two Lists
'''
''' @param {List} Sequence
''' @return {List}
'''
Public Function Concat(ByRef Sequence As List) As List
Attribute Concat.VB_Description = "Concatenates two Lists"
    Dim elem As Variant
    
    Set Concat = New List
    For Each elem In seq_pCollection
        Concat.Add elem
    Next elem
    
    If Sequence Is Nothing Then
        Exit Function
    End If
    
    For Each elem In Sequence
        Concat.Add elem
    Next elem
End Function

'''
''' Returns the index of a last element that equals to Parameter
'''
''' @param {Variant} Item
''' @return {Long}
'''
Public Function LastIndexOf(Item As Variant) As Long
Attribute LastIndexOf.VB_Description = "Returns the index of a last element that equals to Parameter"
    Dim i As Long
    Dim elem As Variant
    
    LastIndexOf = MinIndex - 1
    i = MinIndex
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Item) Then
            LastIndexOf = i
        End If
        
        i = i + 1
    Next elem
End Function

'''
''' Removes an element with specified index from the List
'''
''' @param {Long} Index
'''
Public Sub Remove(ByVal Index As Long)
Attribute Remove.VB_Description = "Removes an element with specified index from the List"
    Dim collIndex As Long
    collIndex = GetCollectionIndex(Index)
        
    seq_pCollection.Remove collIndex
End Sub

'''
''' Inverts the order of the elements in a List
'''
''' @return {List}
'''
Public Function Reverse() As List
Attribute Reverse.VB_Description = "Inverts the order of the elements in a List"
    Dim i As Long
    Set Reverse = New List
    
    For i = seq_pCollection.Count To 1 Step -1
        Reverse.Add seq_pCollection(i)
    Next i
End Function

'''
''' Copies the List
'''
''' @return {List}
'''
Public Function Copy() As List
Attribute Copy.VB_Description = "Copies the List"
    Dim elem As Variant
    
    Set Copy = New List
    For Each elem In seq_pCollection
        Copy.Add elem
    Next elem
End Function

'''
''' Returns a String representing a List
'''
''' @param {String} Delimiter
''' @return {String}
'''
Public Function ToString(Optional ByVal Delimiter As String = " ") As String
Attribute ToString.VB_Description = "Returns a String representing a List"
    Dim elems() As Variant
    Dim elem As Variant
    Dim i As Long
    If seq_pCollection.Count = 0 Then
        Erase elems
        Exit Function
    End If
    
    ReDim elems(1 To seq_pCollection.Count)
    
    i = 1
    For Each elem In seq_pCollection
        If VBA.IsObject(elem) Then
            If elem Is Me Then
                elems(i) = "[THIS LIST]"
            ElseIf elem Is Nothing Then
                elems(i) = "Object{Nothing}"
            ElseIf TypeOf elem Is List Then
                elems(i) = "[" & elem.ToString(Delimiter) & "]"
            Else
                elems(i) = "OBJECT{" & VBA.TypeName(elem) & "}"
            End If
        Else
            elems(i) = CStr(elem)
        End If
    
        i = i + 1
    Next elem
    
    ToString = VBA.Join(elems, Delimiter)
    Erase elems
End Function

'''
''' Returns True if the List if empty
'''
''' @return {Boolean}
'''
Public Function IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Returns True if the List if empty"
    IsEmpty = (seq_pCollection.Count = 0)
End Function

'''
''' Clears the List
'''
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the List"
    Set seq_pCollection = New VBA.Collection
End Sub

'''
''' Determines whether two Lists are equal
'''
''' @param {List} List
''' @return {Boolean}
'''
Public Function Equals(ByRef List As List) As Boolean
Attribute Equals.VB_Description = "Determines whether two Lists are equal"
    Dim i As Long
    Dim elem As Variant
    If List Is Nothing Then
        Equals = False
        Exit Function
    End If
    
    If List.Count <> Me.Count Then
        Equals = False
        Exit Function
    End If
    
    i = 1
    For Each elem In List
        If Not ItemsEqual(elem, seq_pCollection(i)) Then
            Equals = False
            Exit Function
        End If
        
        i = i + 1
    Next elem
    
    Equals = True
End Function

'''
''' Returns a VBA.Collection of Lists elements
'''
''' @return {VBA.Collection}
'''
Public Function ToCollection() As VBA.Collection
Attribute ToCollection.VB_Description = "Returns a Collection of Lists elements"
    Dim elem As Variant
    
    Set ToCollection = New VBA.Collection
    For Each elem In seq_pCollection
        ToCollection.Add elem
    Next elem
End Function

'''
''' Filters a sequence using Predicate
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {List}
'''
Public Function Filter(ByVal Predicate As String, Optional ByRef PredicateOwner As Object = Nothing) As List
Attribute Filter.VB_Description = "Filters a sequence using a Predicate"
    Dim elem As Variant
    Set Filter = New List
    If PredicateOwner Is Nothing Then
        For Each elem In seq_pCollection
            If Application.Run(Predicate, elem) Then
                Filter.Add elem
            End If
        Next elem
    Else
        For Each elem In seq_pCollection
            If VBA.CallByName(PredicateOwner, Predicate, VBA.VbCallType.VbMethod, elem) Then
                Filter.Add elem
            End If
        Next elem
    End If
End Function

'''
''' Projects each element of a sequence using Predicate
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {List}
'''
Public Function Map(ByVal Predicate As String, Optional ByRef PredicateOwner As Object = Nothing) As List
Attribute Map.VB_Description = "Projects each element of a sequence using Predicate"
    Dim elem As Variant
    Set Map = New List
    If PredicateOwner Is Nothing Then
        For Each elem In seq_pCollection
            Map.Add Application.Run(Predicate, elem)
        Next elem
    Else
        For Each elem In seq_pCollection
            Map.Add VBA.CallByName(PredicateOwner, Predicate, VBA.VbCallType.VbMethod, elem)
        Next elem
    End If
End Function

'''
''' Returns distinct elements from the List
'''
''' @return {List}
'''
Public Function Distinct() As List
Attribute Distinct.VB_Description = "Returns distinct elements from the List"
#If UseDictionary Then
    Dim elem As Variant
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    For Each elem In seq_pCollection
        dict(elem) = True
    Next elem
    
    Set Distinct = New List
    For Each elem In dict.Keys
        Distinct.Add elem
    Next elem
    
    Set dict = Nothing
#Else
    Dim coll As VBA.Collection
    Dim elem As Variant
    Dim i As Long
    
    Set coll = New VBA.Collection
    For Each elem In seq_pCollection
        coll.Add elem
    Next elem
    
    Set Distinct = New List
    Do While coll.Count <> 0
        If VBA.IsObject(coll.Item(1)) Then
            Set elem = coll.Item(1)
        Else
            elem = coll.Item(1)
        End If
        
        coll.Remove 1
        Distinct.Add elem
        
        For i = coll.Count To 1 Step -1
            If ItemsEqual(coll.Item(i), elem) Then
                coll.Remove i
            End If
        Next i
    Loop
    
    Set coll = Nothing
#End If
End Function

'''
''' Produces the set intersection of two Lists
'''
''' @param {List} List
''' @return {List}
'''
Public Function Intersect(ByRef List As List) As List
Attribute Intersect.VB_Description = "Produces the set intersection of two Lists"
#If UseDictionary Then
    Dim dict As Dictionary
    Dim elem As Variant
    
    Set dict = New Dictionary
    Set Intersect = New List
    
    If List Is Nothing Then
        Exit Function
    End If
    
    For Each elem In seq_pCollection
        dict(elem) = True
    Next elem
    
    For Each elem In List
        If dict.Exists(elem) Then
            Intersect.Add elem
        End If
    Next elem
    
    Set dict = Nothing
#Else
    Dim elem As Variant
    
    Set Intersect = New List
    If List Is Nothing Then
        Exit Function
    End If
    
    For Each elem In seq_pCollection
        If List.Contains(elem) Then
            Intersect.Add elem
        End If
    Next elem
    
    Set Intersect = Intersect.Distinct()
#End If
End Function

'''
''' Sorts the numeric elements of a List
''' in ascending order
'''
''' @return {List}
'''
Public Function SortNumbers() As List
Attribute SortNumbers.VB_Description = "Sorts the numeric elements of a List in ascending order"
    Dim sz As Long
    Dim elem As Variant
    Dim i As Long
    Dim arr() As Variant
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If VBA.IsNumeric(elem) Then
            arr(i) = CDbl(elem)
            i = i + 1
        End If
    Next elem
    
    sz = i
    QSort Elements:=arr, _
          IndexStart:=0, _
          IndexEnd:=sz, _
          Predicate:=vbNullString, _
          PredicateOwner:=Nothing
    
    Set SortNumbers = New List
    For i = 0 To sz - 1 Step 1
        SortNumbers.Add arr(i)
    Next i
    
    Erase arr
End Function

'''
''' Sorts the elements of a List that can be converted to String
''' in ascending alphabetical order
'''
''' @return {List}
'''
Public Function SortAlphabetically() As List
Attribute SortAlphabetically.VB_Description = "Sorts the elements of a List that can be converted to String in ascending alphabetical order"
    Dim sz As Long
    Dim elem As Variant
    Dim i As Long
    Dim outStr As String
    Dim arr() As Variant
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If TryConvertToString(elem, outStr) Then
            arr(i) = outStr
            i = i + 1
        End If
    Next elem
    
    sz = i
    QSort Elements:=arr, _
          IndexStart:=0, _
          IndexEnd:=sz, _
          Predicate:=vbNullString, _
          PredicateOwner:=Nothing
    
    Set SortAlphabetically = New List
    For i = 0 To sz - 1 Step 1
        SortAlphabetically.Add arr(i)
    Next i
    
    Erase arr
End Function

'''
''' Sorts the elements of a List in ascending order
''' by Predicate or by operators > and < if the Predicate
''' is not specified
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {List}
'''
Public Function Sort(Optional ByVal Predicate As String = vbNullString, Optional ByRef PredicateOwner As Object = Nothing) As List
Attribute Sort.VB_Description = "Sorts the elements of a List in ascending order by Predicate or by operators > and < if the Predicate is not specified"
    Dim arr() As Variant
    Dim elem As Variant
    Dim i As Long
    If seq_pCollection.Count = 0 Then
        Set Sort = New List
        Exit Function
    End If
    
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If VBA.IsObject(elem) Then
            Set arr(i) = elem
        Else
            arr(i) = elem
        End If
        
        i = i + 1
    Next elem
    
    QSort Elements:=arr, _
          IndexStart:=0, _
          IndexEnd:=seq_pCollection.Count, _
          Predicate:=Predicate, _
          PredicateOwner:=PredicateOwner
    
    Set Sort = New List
    For i = 0 To seq_pCollection.Count - 1 Step 1
        Sort.Add arr(i)
    Next i
    
    Erase arr
End Function

'''
''' Sorts the elements of a List in descending order
''' by Predicate or by operators > and < if the Predicate
''' is not specified
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {List}
'''
Public Function SortDescending(Optional ByVal Predicate As String = vbNullString, Optional ByRef PredicateOwner As Object = Nothing) As List
Attribute SortDescending.VB_Description = "Sorts the elements of a List in descending order by Predicate or by operators > and < if the Predicate is not specified"
    Dim arr() As Variant
    Dim elem As Variant
    Dim i As Long
    If seq_pCollection.Count = 0 Then
        Set SortDescending = New List
        Exit Function
    End If
    
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If VBA.IsObject(elem) Then
            Set arr(i) = elem
        Else
            arr(i) = elem
        End If
        
        i = i + 1
    Next elem
    
    QSort Elements:=arr, _
          IndexStart:=0, _
          IndexEnd:=seq_pCollection.Count, _
          Predicate:=Predicate, _
          PredicateOwner:=PredicateOwner
    
    Set SortDescending = New List
    For i = seq_pCollection.Count - 1 To 0 Step -1
        SortDescending.Add arr(i)
    Next i
    
    Erase arr
End Function

'''
''' Sets the element in the List by index
'''
''' @param {Long} Index
''' @param {Variant} Value
'''
Public Sub SetElement(ByVal Index As Long, Value As Variant)
Attribute SetElement.VB_Description = "Sets the element in the List by index"
    Dim collIndex As Long
    collIndex = GetCollectionIndex(Index)

    seq_pCollection.Add Value, After:=collIndex
    seq_pCollection.Remove collIndex
End Sub

'''
''' Swaps positions of two List elements
'''
''' @param {Long} Index1
''' @param {Long} Index2
'''
Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
Attribute Swap.VB_Description = "Swaps positions of two List elements"
    Dim tmp As Variant
    
    If VBA.IsObject(Item(Index1)) Then
        Set tmp = Item(Index1)
    Else
        tmp = Item(Index1)
    End If
    
    SetElement Index1, Item(Index2)
    SetElement Index2, tmp
End Sub

'''
''' Computes the sum of numeric elements of the List
'''
''' @return {Variant}
'''
Public Function Sum() As Variant
Attribute Sum.VB_Description = "Computes the sum of numeric elements of the List"
    Dim elem As Variant
    
    Sum = 0
    For Each elem In seq_pCollection
        If VBA.IsNumeric(elem) Then
            Sum = Sum + elem
        End If
    Next elem
End Function

'''
''' Returns the minimum value
'''
''' @return {Variant}
'''
Public Function Min() As Variant
Attribute Min.VB_Description = "Returns the minimum value"
    Dim elem As Variant
    Min = Empty
    
    For Each elem In seq_pCollection
        If Not VBA.IsObject(elem) Then
            If Min = Empty Then
                Min = elem
            Else
                If Min > elem Then
                    Min = elem
                End If
            End If
        End If
    Next elem
End Function

'''
''' Returns the maximum value
'''
''' @return {Variant}
'''
Public Function Max() As Variant
Attribute Min.VB_Description = "Returns the maximum value"
    Dim elem As Variant
    Max = Empty
    
    For Each elem In seq_pCollection
        If Not VBA.IsObject(elem) Then
            If Max = Empty Then
                Max = elem
            Else
                If Max < elem Then
                    Max = elem
                End If
            End If
        End If
    Next elem
End Function

'''
''' Computes the average of numeric elements of the List
'''
''' @return {Double}
'''
Public Function Average() As Double
Attribute Average.VB_Description = "Computes the average of numeric elements of the List"
    Dim elem As Variant
    Dim cnt As Double
    
    cnt = 0
    For Each elem In seq_pCollection
        If VBA.IsNumeric(elem) Then
            cnt = cnt + 1
        End If
    Next elem
    
    For Each elem In seq_pCollection
        If VBA.IsNumeric(elem) Then
            Average = Average + elem / cnt
        End If
    Next elem
End Function

'''
''' Computes the percentile of numeric elements of the List
'''
''' @param {Double} Percent
''' @return {Double}
'''
Public Function Percentile(ByVal Percent As Double) As Double
Attribute Percentile.VB_Description = "Computes the percentile of numeric elements of the List"
    Dim elem As Variant
    Dim nums As Collection
    Dim r As Double
    Dim ir As Long
    Dim fr As Double
    Dim num1 As Long
    Dim num2 As Long
    If Percent < 0 Or Percent > 1 Then
        Err.Raise Number:=9 'Subscript out of range
        Exit Function
    End If
    
    Set nums = Me.SortNumbers().ToCollection()
    If nums.Count = 0 Then
        Percentile = 0
        Exit Function
    End If

    If nums.Count = 1 Then
        Percentile = nums.Item(1)
        Exit Function
    End If
    
    r = CDbl((0 + nums.Count - 1) * Percent)
    ir = CLng(VBA.Round(r, 0))
    If ir > r Then
        ir = ir - 1
    End If
    
    fr = r - ir
    num1 = nums.Item(ir + 1)
    num2 = nums.Item(ir + 2)
    
    Percentile = num1 + (num2 - num1) * fr
    Set nums = Nothing
End Function

'''
''' Returns number of a first element
'''
''' @return {Long}
'''
Public Function GetMinIndex() As Long
Attribute GetMinIndex.VB_Description = "Returns number of a first element"
    GetMinIndex = MinIndex
End Function

'''
''' Computes the median of numeric elements of the List
'''
''' @return {Double}
'''
Public Function Median() As Double
Attribute Median.VB_Description = "Computes the median of numeric elements of the List"
    Median = Percentile(0.5)
End Function

'''
''' Groups the List by specified key selector predicate function
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {Dictionary}
'''
#If UseDictionary Then
Public Function GroupBy(ByVal Predicate As String, Optional ByRef PredicateOwner As Object = Nothing) As Dictionary
Attribute GroupBy.VB_Description = "Groups the List by specified key selector predicate function"
    Dim elem As Variant
    Dim val As Variant
    Dim grpColl As List
    Set GroupBy = New Dictionary
    
    For Each elem In seq_pCollection
        If PredicateOwner Is Nothing Then
            Set grpColl = GetDictCollectionOrAdd(GroupBy, Application.Run(Predicate, elem))
        Else
            Set grpColl = GetDictCollectionOrAdd(GroupBy, VBA.CallByName(PredicateOwner, Predicate, elem))
        End If
        
        grpColl.Add elem
    Next elem
    
    Set grpColl = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Private Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDictCollectionOrAdd(ByRef Dictionary As Dictionary, Key As Variant) As List
    If Not Dictionary.Exists(Key) Then
        Set Dictionary(Key) = New List
    End If
    
    Set GetDictCollectionOrAdd = Dictionary(Key)
End Function
#End If

Private Function TryConvertToString(Item As Variant, ByRef Result As String) As Boolean
    On Error GoTo errorHandler
    Result = CStr(Item)
    TryConvertToString = True
    Exit Function
errorHandler:
    TryConvertToString = False
End Function

Private Sub QSort(Elements() As Variant, _
                  ByVal IndexStart As Long, _
                  ByVal IndexEnd As Long, _
                  ByVal Predicate As String, _
                  ByRef PredicateOwner As Object)
    
    Dim stk As VBA.Collection
    Dim collStart As Long
    Dim collEnd As Long
    Dim collSepar As Long
    
    Set stk = New VBA.Collection
      
    stk.Add IndexStart
    stk.Add IndexEnd
    
    Do While stk.Count <> 0
        collEnd = stk(stk.Count)
        stk.Remove stk.Count
        collStart = stk(stk.Count)
        stk.Remove stk.Count
        
        If collEnd - collStart < 2 Then GoTo continueLoop
        collSepar = collStart + ((collEnd - collStart) / 2)
        
        Dim l As Long
        Dim h As Long
        Dim piv As Variant
        Dim temp As Variant
        Dim idx As Long
        l = collStart
        h = collEnd - 2
        
        If VBA.IsObject(Elements(collSepar)) Then
            Set piv = Elements(collSepar)
        Else
            piv = Elements(collSepar)
        End If
        
        SwapArrayElements Elements, collSepar, collEnd - 1
        
        Do While l < h
            If CallCompareFunction(Predicate, PredicateOwner, Elements(l), piv) < 0 Then
                l = l + 1
            ElseIf CallCompareFunction(Predicate, PredicateOwner, Elements(h), piv) >= 0 Then
                h = h - 1
            Else
                SwapArrayElements Elements, l, h
            End If
        Loop
        
        idx = h
        If CallCompareFunction(Predicate, PredicateOwner, Elements(h), piv) < 0 Then
            idx = idx + 1
        End If
        
        SwapArrayElements Elements, collEnd - 1, idx
        
        stk.Add idx + 1
        stk.Add collEnd
        
        stk.Add collStart
        stk.Add idx
        
continueLoop:
        If VBA.IsObject(piv) Then
            Set piv = Nothing
        End If
    Loop
    
    Set stk = Nothing
End Sub

Private Sub SwapArrayElements(Elements() As Variant, ByVal Item1 As Long, ByVal Item2 As Long)
    Dim temp As Variant
    If VBA.IsObject(Elements(Item1)) Then
        Set temp = Elements(Item1)
    Else
        temp = Elements(Item1)
    End If
    
    If VBA.IsObject(Elements(Item2)) Then
        Set Elements(Item1) = Elements(Item2)
    Else
        Elements(Item1) = Elements(Item2)
    End If
    
    If VBA.IsObject(temp) Then
        Set Elements(Item2) = temp
    Else
        Elements(Item2) = temp
    End If
    
    If VBA.IsObject(temp) Then
        Set temp = Nothing
    End If
End Sub

Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = seq_pCollection.[_NewEnum]
End Property

Private Function CallCompareFunction(ByVal FunctionName As String, _
                                     ByRef FunctionOwner As Object, _
                                     Arg1 As Variant, _
                                     Arg2 As Variant) As Variant

    If FunctionName = vbNullString Then
        If Arg1 > Arg2 Then
            CallCompareFunction = 1
        ElseIf Arg1 < Arg2 Then
            CallCompareFunction = -1
        Else
            CallCompareFunction = 0
        End If
    ElseIf FunctionOwner Is Nothing Then
        CallCompareFunction = Application.Run(FunctionName, Arg1, Arg2)
    Else
        CallCompareFunction = VBA.CallByName(FunctionOwner, FunctionName, VbMethod, Arg1, Arg2)
    End If
End Function

Private Function ItemsEqual(Item1 As Variant, Item2 As Variant) As Boolean
    Dim item1Obj As Boolean
    Dim item2Obj As Boolean
    
    item1Obj = VBA.IsObject(Item1)
    item2Obj = VBA.IsObject(Item2)
    
    If item1Obj And item2Obj Then
        ItemsEqual = Item1 Is Item2
    ElseIf (Not item1Obj) And (Not item2Obj) Then
        ItemsEqual = (Item1 = Item2)
    Else
        ItemsEqual = False
    End If
End Function

Private Function GetCollectionIndex(ByVal Index As Long) As Long
    GetCollectionIndex = Index - MinIndex + 1
End Function

Private Sub Class_Initialize()
    Set seq_pCollection = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set seq_pCollection = Nothing
End Sub
