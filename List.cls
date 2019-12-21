VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a list of elements that can be accessed by index. Provides methods to search, sort, and manipulate lists. \r\n\r\nAuthor: Andrew Tokarskiy tokarskiy.a@gmail.com \r\nGithub: https://github.com/tokarskiy \r\nLicense: http://www.opensource.org/licenses/mit-license.php"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' List v0.1
'''
''' Represents a list of elements that can be accessed by index.
''' Provides methods to search, sort and manipulate lists.
'''
''' Author : Andrew Tokarskiy tokarskiy.a@gmail.com
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
''' @param {Variant} Element
''' @return {Long}
'''
Public Function CountOf(Element As Variant) As Long
Attribute CountOf.VB_Description = "Returns the number that represents how many elements equal to the function parameter"
    Dim elem As Variant
    
    CountOf = 0
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Element) Then
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
''' Inserts an element to the List
'''
''' @param {Variant} Element
''' @param {Long} Before
''' @param {Long} After
'''
Public Sub Add(Element As Variant, Optional Before As Long = MinIndex - 1, Optional After As Long = MinIndex - 1)
Attribute Add.VB_Description = "Inserts an element to the List"
    Dim collIndex As Long
    If After >= MinIndex Then
        collIndex = GetCollectionIndex(After)
        seq_pCollection.Add Element, After:=collIndex
    ElseIf Before >= MinIndex Then
        collIndex = GetCollectionIndex(Before)
        seq_pCollection.Add Element, Before:=collIndex
    Else
        seq_pCollection.Add Element
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
''' @param {Excel.Range} Sequence
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
''' @param {Variant} Sequence
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
''' @param {Variant} Element
''' @return {Boolean}
'''
Public Function Contains(Element As Variant) As Boolean
Attribute Contains.VB_Description = "Determines whether any element of a List equals the Parameter"
    Contains = IndexOf(Element) >= MinIndex
End Function

'''
''' Returns the index of a first element that equals to Parameter
'''
''' @param {Element} Variant
''' @return {Long}
'''
Public Function IndexOf(Element As Variant) As Long
Attribute IndexOf.VB_Description = "Returns the index of a first element that equals to Parameter"
    Dim i As Long
    Dim elem As Variant
    
    i = MinIndex
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Element) Then
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
''' @param {Element} Variant
''' @return {Long}
'''
Public Function LastIndexOf(Element As Variant) As Long
Attribute LastIndexOf.VB_Description = "Returns the index of a last element that equals to Parameter"
    Dim i As Long
    Dim elem As Variant
    
    LastIndexOf = MinIndex - 1
    i = MinIndex
    For Each elem In seq_pCollection
        If ItemsEqual(elem, Element) Then
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
        If IsObject(elem) Then
            If elem Is Me Then
                elems(i) = "[THIS LIST]"
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
    For Each elem In seq_pCollection
        If CallFunction1(Predicate, PredicateOwner, elem) Then
            Filter.Add elem
        End If
    Next elem
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
    For Each elem In seq_pCollection
        Map.Add CallFunction1(Predicate, PredicateOwner, elem)
    Next elem
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
        If IsObject(coll.Item(1)) Then
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
    
    For Each elem In List
        If Me.Contains(elem) Then
            Intersect.Add elem
        End If
    Next elem
    
    Set Intersect = Intersect.Distinct()
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
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If IsObject(elem) Then
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
    ReDim arr(0 To seq_pCollection.Count - 1)
    
    i = 0
    For Each elem In seq_pCollection
        If IsObject(elem) Then
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
''' Groups the List by specified key selector predicate function
'''
''' @param {String} Predicate
''' @param {Object} PredicateOwner
''' @return {Dictionary}
'''
#If UseDictionary Then
Public Function GroupBy(Optional ByVal Predicate As String = vbNullString, Optional ByRef PredicateOwner As Object = Nothing) As Dictionary
Attribute GroupBy.VB_Description = "Groups the List by specified key selector predicate function"
    Dim elem As Variant
    Dim val As Variant
    Dim grpColl As List
    Set GroupBy = New Dictionary
    
    For Each elem In seq_pCollection
        Set grpColl = GetDictCollectionOrAdd(GroupBy, CallFunction1(Predicate, PredicateOwner, elem))
        
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
      
    stk.Add 0
    stk.Add seq_pCollection.Count
    
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
        
        If IsObject(Elements(collSepar)) Then
            Set piv = Elements(collSepar)
        Else
            piv = Elements(collSepar)
        End If
        
        SwapArrayElements Elements, collSepar, collEnd - 1
        
        Do While l < h
            If CallFunction2(Predicate, PredicateOwner, Elements(l), piv) < 0 Then
                l = l + 1
            ElseIf CallFunction2(Predicate, PredicateOwner, Elements(h), piv) >= 0 Then
                h = h - 1
            Else
                SwapArrayElements Elements, l, h
            End If
        Loop
        
        idx = h
        If CallFunction2(Predicate, PredicateOwner, Elements(h), piv) < 0 Then
            idx = idx + 1
        End If
        
        SwapArrayElements Elements, collEnd - 1, idx
        
        stk.Add idx + 1
        stk.Add collEnd
        
        stk.Add collStart
        stk.Add idx
        
continueLoop:
    Loop
    
    Set stk = Nothing
End Sub

Private Sub SwapArrayElements(Elements() As Variant, ByVal Item1 As Long, ByVal Item2 As Long)
    Dim temp As Variant
    If IsObject(Elements(Item1)) Then
        Set temp = Elements(Item1)
    Else
        temp = Elements(Item1)
    End If
    
    If IsObject(Elements(Item2)) Then
        Set Elements(Item1) = Elements(Item2)
    Else
        Elements(Item1) = Elements(Item2)
    End If
    
    If IsObject(temp) Then
        Set Elements(Item2) = temp
    Else
        Elements(Item2) = temp
    End If
End Sub

Private Function CallFunction1(ByVal FunctionName As String, ByRef FunctionOwner As Object, Arg1 As Variant) As Variant
    If FunctionName = vbNullString Then
        CallFunction1 = True
    ElseIf FunctionOwner Is Nothing Then
        CallFunction1 = Application.Run(FunctionName, Arg1)
    Else
        CallFunction1 = VBA.CallByName(FunctionOwner, FunctionName, VbMethod, Arg1)
    End If
End Function

Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = seq_pCollection.[_NewEnum]
End Property

Private Function CallFunction2(ByVal FunctionName As String, ByRef FunctionOwner As Object, Arg1 As Variant, Arg2 As Variant) As Variant
    If FunctionName = vbNullString Then
        If Arg1 > Arg2 Then
            CallFunction2 = 1
        ElseIf Arg1 < Arg2 Then
            CallFunction2 = -1
        Else
            CallFunction2 = 0
        End If
    ElseIf FunctionOwner Is Nothing Then
        CallFunction2 = Application.Run(FunctionName, Arg1, Arg2)
    Else
        CallFunction2 = VBA.CallByName(FunctionOwner, FunctionName, VbMethod, Arg1, Arg2)
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
