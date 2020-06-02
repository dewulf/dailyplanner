'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 11/22/2013 11:15:17 AM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414365/raw/usefulcJobject.vba
'v2.5
Option Explicit

Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject
    Dim j As New cJobject
    Set JSONParse = j.init(Nothing).parse(s, jtype, complain)
    j.tearDown
End Function
Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String
    JSONStringify = j.stringify(blf)
End Function
Public Function jSonArgs(options As String) As cJobject
    ' takes a javaScript like options paramte and converts it to cJobject
    ' it can be accessed as job.child('argName').value or job.find('argName') etc.
    Dim job As New cJobject
    If options <> vbNullString Then
        Set jSonArgs = job.init(Nothing, "jSonArgs").deSerialize(options)
    End If
End Function
Public Function optionsExtend(givenOptions As String, _
            Optional defaultOptions As String = vbNullString) As cJobject
    Dim jGiven As cJobject, jDefault As cJobject, _
        jExtended As cJobject, cj As cJobject
    ' this works like $.extend in jQuery.
    ' given and default options arrive as a json string
    ' example -
    ' optionsExtend ("{'width':90,'color':'blue'}", "{'width':20,'height':30,'color':'red'}")
    ' would return a cJobject which serializes to
    ' "{width:90,height:30,color:blue}"
    Set jGiven = jSonArgs(givenOptions)
    Set jDefault = jSonArgs(defaultOptions)
    
    ' now we combine them
    If Not jDefault Is Nothing Then
        Set jExtended = jDefault
    Else
        Set jExtended = New cJobject
        jExtended.init Nothing
    End If
    
    ' now we merge that with whatever was given
    If Not jGiven Is Nothing Then
        jExtended.merge jGiven
    End If
    
    ' and its over
    Set optionsExtend = jExtended
End Function

'udfs to expose classes
Public Function ucJobjectMake(r As Variant) As cJobject
    Dim cj As New cJobject
    Set ucJobjectMake = cj.deSerialize(CStr(r))
End Function
Public Function ucJobjectChildValue(json As Variant, child As Variant) As String
    ucJobjectChildValue = ucJobjectMake(CStr(json)).child(CStr(child)).value
End Function
Public Function ucJobjectLint(json As Variant, Optional child As Variant) As String
    Dim cj As cJobject
    Set cj = ucJobjectMake(json)
    If Not IsMissing(child) Then
        Set cj = cj.child(CStr(child))
    End If
    ucJobjectLint = cj.serialize(True)
End Function
Public Function cleanGoogleWire(sWire As String) As String
    Dim jStart As String, p As Long, newWire As Boolean, e As Long, s As String
    
    jStart = "table:"
    p = InStr(1, sWire, jStart)
    'there have been multiple versions of wire ...
    If p = 0 Then
        'try the other one
        jStart = q & ("table") & q & ":"
        p = InStr(1, sWire, jStart)
        newWire = True
    End If

    p = InStr(1, sWire, jStart)
    e = Len(sWire) - 1

    If p <= 0 Or e <= 0 Or p > e Then
        MsgBox " did not find table definition data"
        Exit Function
    End If
    
    If Mid(sWire, e, 2) <> ");" Then
        MsgBox ("incomplete google wire message")
        Exit Function
    End If
    ' encode the 'table:' part to a cjobject
    p = p + Len(jStart)
    s = "{" & jStart & "[" & Mid(sWire, p, e - p - 1) & "]}"
    ' google protocol doesnt have quotes round the key of key value pairs,
    ' and i also need to convert date from javascript syntax new Date()
    s = rxReplace("(new\sDate)(\()(\d+)(,)(\d+)(,)(\d+)(\))", s, "'$3/$5/$7'")
    If Not newWire Then s = rxReplace("(\w+)(:)", s, "'$1':")
    cleanGoogleWire = s
    
End Function

Public Function xmlStringToJobject(xmlString As String, Optional complain As Boolean = True) As cJobject
    Dim doc As Object
    ' parse xml

    Set doc = CreateObject("msxml2.DOMDocument")
    doc.LoadXML xmlString
    If doc.parsed And doc.parseError = 0 Then
        Set xmlStringToJobject = docToJobject(doc, complain)
        Exit Function
    End If

    Set xmlStringToJobject = Nothing
    If complain Then
        MsgBox ("Invalid xml string - xmlparseerror code:" & doc.parseError)
    End If
    
    Exit Function
    
End Function
Public Function docToJobject(doc As Object, Optional complain As Boolean = True) As cJobject
    ' convert xml document to a cjobject
    Dim node As IXMLDOMNode, job As cJobject
    Set job = New cJobject
    job.init Nothing
       
    Set docToJobject = handleNodes(doc, job)
End Function
Private Function isArrayRoot(parent As IXMLDOMNode) As Boolean
    
    Dim node As IXMLDOMNode, n As Long, node2 As IXMLDOMNode
    
    
    isArrayRoot = False
    If parent.NodeType = NODE_ELEMENT And parent.ChildNodes.Length > 1 Then
        For Each node2 In parent.ChildNodes
            If node2.NodeType = NODE_ELEMENT Then
                n = 0
                For Each node In parent.ChildNodes
                    If node.NodeType = NODE_ELEMENT And _
                        node2.nodeName = node.nodeName Then n = n + 1
                Next node
                If n > 1 Then
                    ' this shoudl be true, but for leniency i'll comment
                    'Debug.Assert n = parent.ChildNodes.Length
                    isArrayRoot = True
                    Exit Function
                End If
            End If
        Next node2
    End If

    
End Function
Private Function handleNodes(parent As IXMLDOMNode, job As cJobject) As cJobject
    Dim node As IXMLDOMNode, joc As cJobject, attrib As IXMLDOMAttribute, i As Long, _
         arrayJob As cJobject
    
    If isArrayRoot(parent) Then
        ' we need an array associated with this this node
        ' subsequent members will need to make space for themselves
        Set joc = job.add(parent.nodeName).addArray
    Else
        Set joc = handleNode(parent, job)
    End If
    
    ' deal with any attributes
    If Not parent.Attributes Is Nothing Then
        For Each attrib In parent.Attributes
            handleNode attrib, joc
        Next attrib
    End If
    
    ' do the children
    If Not parent.ChildNodes Is Nothing And parent.ChildNodes.Length > 0 Then
        For Each node In parent.ChildNodes
            handleNodes node, joc
        Next node
    End If
    
    ' always return the level at which we arrived
    Set handleNodes = job
    
End Function
Private Function handleNode(node As IXMLDOMNode, job As cJobject, Optional arrayHead As Boolean = False) As cJobject
    Dim key As cJobject

    Set handleNode = job
   
    
    Select Case node.NodeType
        Case NODE_ATTRIBUTE
            ' we cant have an array of attributes - this will silently use the latest
            job.add node.nodeName, node.NodeValue
            
        Case NODE_ELEMENT
            If job.isArrayRoot Then
                If (node.ChildNodes.Length = 1 And _
                        node.ChildNodes(0).NodeType = NODE_TEXT) Then
                    Set handleNode = job.add.add
                Else
                    Set handleNode = job.add.add(node.nodeName)
                End If
            Else
                Set handleNode = job.add(node.nodeName)
            End If

        Case NODE_TEXT
            job.value = node.NodeValue

            
        Case NODE_DOCUMENT, NODE_CDATA_SECTION, NODE_ENTITY_REFERENCE, _
            NODE_ENTITY, NODE_PROCESSING_INSTRUCTION, NODE_COMMENT, NODE_DOCUMENT_TYPE, _
            NODE_DOCUMENT_FRAGMENT, NODE_NOTATION
            ' just ignore these for now

            
        Case Else
            Debug.Assert False
    End Select
    
End Function

