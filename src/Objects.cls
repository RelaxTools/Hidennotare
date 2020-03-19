VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [Hidennotare] v2.5
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/Hidennotare
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------
'>### Objects クラス
'>
'>**Remarks**
'>
'>Objectsクラス<br>
'>オブジェクトの操作に関係する処理をまとめたクラスです。<br>
'>**このクラスは Staticクラス(Attribute VB_PredeclaredId = True) です。**
'>
'>**Dependent**
'>
'>* None
'>
'>**History**
'>
'>|Date|Version|Note|
'>|---|---|---|
'>|2019-01-01|Ver.1.0.0|新規作成|
'>
'-----------------------------------------------------------------------------------------------------
Option Explicit
'>---
'>#### SetWith(Objects)
'>
'>**Syntax**
'>
'>```
'>Objects.SetWith(inObj) = obj
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|Variant|代入先のオブジェクトまたは変数|
'>obj|必須|Variant|代入元のオブジェクトまたは変数|
'>
'>**Return Value**
'>
'>VBAではプリミティブ型とオブジェクト型の変数の設定方法が異なる。<br>
'>プリミティブ型では Let オブジェクト型なら Set をつけなければならないが<br>
'>プロパティ内で隠蔽可能。<br>
'>
'>**Remarks**
'>
'>Set/Letプロパティ<br>
'>VBA 個人的汎用処理 https://qiita.com/nukie_53/items/bde16afd9a6ca789949d<br>
'>@nukie_53<br>
'>Set/Letを隠蔽するプロパティ<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IValueObject
'>
Public Property Let SetWith(outVariable As Variant, inExpression As Variant)
    
    Select Case True
        Case VBA.IsObject(inExpression), VBA.VarType(inExpression) = vbDataObject
            Set outVariable = inExpression
        Case Else
            Let outVariable = inExpression
    End Select

End Property
'>---
'>#### AsIValueObject(Objects)
'>
'>**Syntax**
'>
'>```
'>Set obj = Objects.AsIValueObject(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|IValueObjectに対応したオブジェクト|
'>
'>**Return Value**
'>
'>IValueObjectにキャストされたオブジェクト
'>
'>**Remarks**
'>
'>IValueObject変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IValueObject
'>
Public Function AsIValueObject(ByVal obj As Object) As IValueObject
    Set AsIValueObject = obj
End Function
'>---
'>#### AsINewInstance(Objects)
'>
'>**Syntax**
'>
'>```
'>Set obj = Objects.AsINewInstance(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|INewInstanceに対応したオブジェクト|
'>
'>**Return Value**
'>
'>INewInstanceにキャストされたオブジェクト
'>
'>**Remarks**
'>
'>INewInstance変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* INewInstance
'>
Public Function AsINewInstance(ByVal obj As Object) As INewInstance
    Set AsINewInstance = obj
End Function
'>---
'>#### AsIList(Objects)
'>
'>**Syntax**
'>
'>```
'>Set obj = Objects.AsIList(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|IListに対応したオブジェクト|
'>
'>**Return Value**
'>
'>IListにキャストされたオブジェクト
'>
'>**Remarks**
'>
'>IList変換<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IList
'>
Public Function AsIList(ByVal obj As Object) As IList
    Set AsIList = obj
End Function
'>---
'>#### AsIDictionary(Objects)
'>
'>**Syntax**
'>
'>```
'>Set obj = Objects.AsIDictionary(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|IDictionaryに対応したオブジェクト|
'>
'>**Return Value**
'>
'>IDictionaryにキャストされたオブジェクト
'>
'>**Remarks**
'>
'>IDictionaryキャスト<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Public Function AsIDictionary(ByVal obj As Object) As IDictionary
    Set AsIDictionary = obj
End Function
'>---
'>#### InstanceOfIDictionary(Objects)
'>
'>**Syntax**
'>
'>```
'>ret = Objects.InstanceOfIDictionary(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|Boolean|検査するオブジェクトト|
'>
'>**Return Value**
'>
'>True:対応/False:未対応
'>
'>**Remarks**
'>
'>IDictionary I/F に対応するクラスか判定する<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Public Function InstanceOfIDictionary(v As Variant) As Boolean

    Dim obj As IDictionary
    
    On Error GoTo e
    
    Set obj = v
    
    InstanceOfIDictionary = True
    Exit Function
e:
    InstanceOfIDictionary = False
    
End Function
'>---
'>#### InstanceOfIList(Objects)
'>
'>**Syntax**
'>
'>```
'>ret = Objects.InstanceOfIList(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|Boolean|検査するオブジェクトト|
'>
'>**Return Value**
'>
'>True:対応/False:未対応
'>
'>**Remarks**
'>
'>IList I/F に対応するクラスか判定する<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Public Function InstanceOfIList(v As Variant) As Boolean

    Dim obj As IList
    
    On Error GoTo e
    
    Set obj = v
    
    InstanceOfIList = True
    Exit Function
e:
    InstanceOfIList = False

End Function
'>---
'>#### InstanceOfIValueObject(Objects)
'>
'>**Syntax**
'>
'>```
'>ret = Objects.InstanceOfIValueObject(inObj)
'>```
'>
'>**Parameters**
'>
'>|Name|Required/Optional|Data type|Description|
'>---|---|---|---
'>inObj|必須|Boolean|検査するオブジェクト|
'>
'>**Return Value**
'>
'>True:対応/False:未対応
'>
'>**Remarks**
'>
'>IValueObject I/F に対応するクラスか判定する<br>
'>
'>**Example**
'>
'>* None
'>
'>**See also**
'>
'>* IDictionary
'>
Public Function InstanceOfIValueObject(v As Variant) As Boolean

    Dim obj As IValueObject
    
    On Error GoTo e
    
    Set obj = v
    
    InstanceOfIValueObject = True
    Exit Function
e:
    InstanceOfIValueObject = False

End Function
