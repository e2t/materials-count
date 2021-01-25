Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim gCurDoc As ModelDoc2
Dim gCurDocName As String
Dim gCurConf As String
Dim gMaterials As Dictionary
Dim gKeys() As String
Dim gFSO As FileSystemObject
Dim gCurDirMask As String
Dim gWherePartIsUsed As Dictionary

Const COL_NAME = 0
Const COL_MASS = 1
Const COL_COUNT = 2

Sub Main()
    Set swApp = Application.SldWorks
    Set gFSO = New FileSystemObject
    Set gMaterials = New Dictionary
    Set gWherePartIsUsed = New Dictionary
    
    Set gCurDoc = swApp.ActiveDoc
    If gCurDoc Is Nothing Then Exit Sub
    If gCurDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Запускать в сборках!", vbCritical
        Exit Sub
    End If
   
    gCurDirMask = LCase(gFSO.GetParentFolderName(gCurDoc.GetPathName)) & "\*"
    gCurDocName = gFSO.GetBaseName(gCurDoc.GetPathName)
    gCurConf = gCurDoc.GetActiveConfiguration.Name
    
    ResearchMaterials
    
    MainForm.Caption = "Материалы " & gCurDocName & " (" & gCurConf & ") "
    MainForm.Show
End Sub

Function ResearchMaterials() 'mask for button
    Dim onlyInCurrentDir As Boolean
    
    onlyInCurrentDir = MainForm.chkCurDir.Value
    gMaterials.RemoveAll
    SearchMaterials gCurDoc, onlyInCurrentDir
    If gMaterials.count > 0 Then
        ReDim gKeys(gMaterials.count - 1)
    End If
    FilterAndPrint
    ClearPartList
End Function

Sub SearchMaterials(asm As AssemblyDoc, onlyInCurrentDir As Boolean)
  Const TopLevelOnly = True
  Dim comp_ As Variant
  Dim comp As Component2
  Dim doc As ModelDoc2
  
  If asm.GetComponentCount(TopLevelOnly) > 0 Then
    For Each comp_ In asm.GetComponents(TopLevelOnly)
      Set comp = comp_
      If comp.IsSuppressed Then  'погашен
        GoTo NextFor
      End If
      Set doc = comp.GetModelDoc2
      If doc Is Nothing Then  'не найден
        GoTo NextFor
      End If
      If Not onlyInCurrentDir Or (LCase(doc.GetPathName) Like gCurDirMask) Then
        If doc.GetType = swDocASSEMBLY Then
          SearchMaterials doc, onlyInCurrentDir
        Else  'doc is part
          AddComponent comp
        End If
      End If
NextFor:
    Next
  End If
End Sub

Sub AddComponent(comp As Component2)
    Dim partName As String
    Dim conf As String
    Dim key As String
    Dim item As MaterialBlankInfo
    Dim material As String
    Dim propBlank As String
    Dim propSize As String
    Dim doc As ModelDoc2
    Dim part As PartDoc
    Dim materialDB As String
    Dim mass As Double
    Dim materialBlank As String
        
    partName = gFSO.GetBaseName(comp.GetPathName)
    conf = comp.ReferencedConfiguration
    Set doc = comp.GetModelDoc2
    Set part = doc
    material = Trim(part.GetMaterialPropertyName2(conf, materialDB))
    propBlank = Trim(GetProp(doc, conf, "Заготовка"))
    propSize = Trim(GetProp(doc, conf, "Типоразмер"))
    mass = GetMass(doc, comp.GetBody)
    
    materialBlank = CreateMaterialBlank(material, propBlank, propSize)
    key = LCase(materialBlank)
    
    If Not gMaterials.Exists(key) Then
        Set item = New MaterialBlankInfo
        item.materialBlank = materialBlank
        item.totalCount = 0
        item.totalMass = 0#
        Set item.where = New Dictionary
        gMaterials.Add key, item
    End If
        
    AddWherePartIsUsed key, partName, conf, mass, doc
    gMaterials(key).totalCount = gMaterials(key).totalCount + 1
    gMaterials(key).totalMass = gMaterials(key).totalMass + mass
End Sub

Function GetProp(doc As ModelDoc2, conf As String, prop As String) As String
    Dim rawValue As String
    Dim resolvedValue As String
    Dim wasResolved As Boolean
    
    resolvedValue = ""
    If doc.Extension.CustomPropertyManager(conf).Get5(prop, False, rawValue, resolvedValue, wasResolved) = swCustomInfoGetResult_NotPresent Then
        doc.Extension.CustomPropertyManager("").Get5 prop, False, rawValue, resolvedValue, wasResolved
    End If
    GetProp = resolvedValue
End Function

Function GetMass(doc As ModelDoc2, body As Body2) As Double
    Dim massProperties() As Double
    Dim result As swMassPropertiesStatus_e
    Dim density As Double

    GetMass = 0#
    If body Is Nothing Then
        'Пустые детали, например, клей, герметик.
        Exit Function
    End If
    density = doc.GetUserPreferenceDoubleValue(swMaterialPropertyDensity)
    massProperties = body.GetMassProperties(density)
    If result = swMassPropertiesStatus_OK Then
        GetMass = massProperties(5)
    End If
End Function

Function CreateMaterialBlank(material As String, propBlank As String, propSize As String) As String
    CreateMaterialBlank = material
    If propBlank <> "" Or propSize <> "" Then
        CreateMaterialBlank = CreateMaterialBlank & ","
    End If
    If propBlank <> "" Then
        CreateMaterialBlank = CreateMaterialBlank & " " & propBlank
    End If
    If propSize <> "" Then
        CreateMaterialBlank = CreateMaterialBlank & " " & propSize
    End If
End Function

Sub AddWherePartIsUsed(key As String, partName As String, conf As String, mass As Double, doc As ModelDoc2)
    Dim partKey As String
    Dim item As WhereInfo
    
    partKey = LCase(partName & "/*@@*/" & conf)
    If Not gMaterials(key).where.Exists(partKey) Then
        Set item = New WhereInfo
        item.partName = partName
        item.conf = conf
        item.count = 0
        item.mass = 0#
        Set item.doc = doc
        gMaterials(key).where.Add partKey, item
    End If
    
    gMaterials(key).where(partKey).count = gMaterials(key).where(partKey).count + 1
    gMaterials(key).where(partKey).mass = gMaterials(key).where(partKey).mass + mass
End Sub

Function FilterAndPrint() 'mask for button
    Dim key_ As Variant
    Dim index As Long
    Dim info As MaterialBlankInfo
    Dim masks As Variant
    
    index = -1
    If gMaterials.count > 0 Then
        masks = CreateFilters
        For Each key_ In gMaterials.keys
            Set info = gMaterials(key_)
            If CheckUserFilter(info.materialBlank, masks) Then
                index = index + 1
                gKeys(index) = key_
            End If
        Next
        If index >= 0 Then
            QuickSort gKeys, LBound(gKeys), index
        End If
    End If
    PrintComponents index
End Function

Function CheckUserFilter(baseName As String, masks As Variant) As Boolean
    Dim i As Integer
    
    CheckUserFilter = True
    If Not IsArrayEmpty(masks) Then
        For i = LBound(masks) To UBound(masks)
            If Not LCase(baseName) Like masks(i) Then
                CheckUserFilter = False
                Exit Function
            End If
        Next
    End If
End Function

Function CreateFilters() As String()
    Dim filter As String
    Dim words As Variant
    Dim i As Integer
    
    filter = MainForm.txtFilter.Value
    words = Split(filter, " ")
    If Not IsArrayEmpty(words) Then
        For i = LBound(words) To UBound(words)
            words(i) = LCase("*" & words(i) & "*")
        Next
    End If
    CreateFilters = words
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean
    Dim i As Integer
  
    On Error GoTo ArrayIsEmpty
    IsArrayEmpty = LBound(anArray) > UBound(anArray)
    Exit Function
ArrayIsEmpty:
    IsArrayEmpty = True
End Function

Function PrintComponents(topKeysBound As Long) 'mask for button
    Dim info As MaterialBlankInfo
    Dim i As Integer
    
    With MainForm.lstDeps
        .Clear
        For i = 0 To topKeysBound
            Set info = gMaterials(gKeys(i))
                .AddItem
                .List(.ListCount - 1, COL_NAME) = info.materialBlank
                .List(.ListCount - 1, COL_MASS) = Format(info.totalMass, "0.###") & " кг"
                .List(.ListCount - 1, COL_COUNT) = Str(info.totalCount) & " шт."
        Next
    End With
End Function

Function ClearPartList() 'mask for button
    MainForm.lstParts.Clear
    gWherePartIsUsed.RemoveAll
End Function

Function KeyForWhereIsPartUsed(where As WhereInfo) As String
    KeyForWhereIsPartUsed = where.partName & " (" & where.conf & ")"
End Function

Sub ShowWhereIsPartUsed(index As Integer)
    Dim key As String
    Dim info As MaterialBlankInfo
    Dim partKey_ As Variant
    Dim parts() As String
    Dim i As Long
    Dim where As WhereInfo
    
    ClearPartList
    key = gKeys(index)
    Set info = gMaterials(key)
    ReDim parts(info.where.count)
    i = -1
    For Each partKey_ In info.where
        i = i + 1
        Set where = info.where(partKey_)
        parts(i) = KeyForWhereIsPartUsed(where) & " --" _
                   & Str(where.count) & " шт." _
                   '& "[" & Str(where.mass) & " kg ]"
        gWherePartIsUsed.Add parts(i), where
    Next
    QuickSort parts, LBound(parts), i
    For Each partKey_ In parts
        MainForm.lstParts.AddItem partKey_
    Next
End Sub

Sub OpenPart(index As Integer, needToCopy As Boolean)
    Dim errors As swActivateDocError_e
    Dim key As String
    Dim where As WhereInfo
    
    key = MainForm.lstParts.List(index)
    Set where = gWherePartIsUsed(key)
    swApp.ActivateDoc3 where.doc.GetPathName, False, swDontRebuildActiveDoc, errors
    where.doc.ShowConfiguration2 where.conf
    
    If needToCopy Then
        CopyInClipBoard CreateNewName(where.doc.Extension, where.conf)
    End If
End Sub

Function GetProperty(propName As String, docext As ModelDocExtension, confName As String) As String
    Dim resultGetProp As swCustomInfoGetResult_e
    Dim rawProp As String, resolvedValue As String
    Dim wasResolved As Boolean
    
    resultGetProp = docext.CustomPropertyManager(confName).Get5(propName, True, rawProp, resolvedValue, wasResolved)
    If resultGetProp = swCustomInfoGetResult_NotPresent Then
        docext.CustomPropertyManager("").Get5 propName, True, rawProp, resolvedValue, wasResolved
    End If
    GetProperty = resolvedValue
End Function

Function CreateNewName(docext As ModelDocExtension, confName As String) As String
    CreateNewName = GetProperty("Обозначение", docext, confName) + " " + _
                    GetProperty("Наименование", docext, confName)
End Function

Function ExitApp()  'mask for button
    Unload MainForm
    End
End Function
