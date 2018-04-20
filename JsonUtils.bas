Attribute VB_Name = "JsonUtils"
Option Explicit

Public Function convertSlideToJson(ByRef slide As PowerPoint.slide, ByRef params As JsonConverterParams) As String
    
    Dim jsonObject As New Dictionary
    Dim copyShapes As New Collection
    Dim shapeInfo As Dictionary
    'Find the center because that is who we flip around
    Dim shape As PowerPoint.shape
    Dim centersShape As PowerPoint.shape
    For Each shape In slide.shapes
        If shape.HasTextFrame Then
            If StrComp(shape.TextFrame.TextRange.Text, "C", vbTextCompare) = 0 Then
                Set centersShape = shape
                Exit For
            End If
        End If
    Next shape
    
    For Each shape In slide.shapes
        If shape.Type = msoAutoShape And Not params.IgnorePlayers Then
            Dim textInShape As String
            textInShape = shape.TextFrame.TextRange.Text
            
            'if we are ignoring lineman, only proceed if the shape does not contain text G, C, or T
            If Not params.IgnoreLineman Or _
                    (StrComp(textInShape, "G", vbTextCompare) <> 0 And StrComp(textInShape, "C", vbTextCompare) And StrComp(textInShape, "T", vbTextCompare)) Then
                Set shapeInfo = New Dictionary
                If shape.AutoShapeType = msoShapeRectangle Then
                    shapeInfo.add "Type", "Rectangle"
                Else
                    shapeInfo.add "Type", "Oval"
                End If
                shapeInfo.add "Label", shape.TextFrame.TextRange.Text
                shapeInfo.add "Label Size", shape.TextFrame.TextRange.Font.Size
                If params.Flip Then
                    If shape.left < centersShape.left Then
                        shapeInfo.add "Left", centersShape.left + centersShape.width + (centersShape.left - (shape.left + shape.width))
                    Else
                        shapeInfo.add "Left", centersShape.left - (shape.left - (centersShape.left + centersShape.width)) - shape.width
                    End If
                Else
                    shapeInfo.add "Left", shape.left
                End If
                shapeInfo.add "Top", shape.top
                shapeInfo.add "Width", shape.width
                shapeInfo.add "Height", shape.height
                shapeInfo.add "Fill Color", shape.Fill.ForeColor.RGB
                shapeInfo.add "Text Color", shape.TextFrame.TextRange.Font.Color.RGB
                shapeInfo.add "Outline Color", shape.Line.ForeColor.RGB
                copyShapes.add shapeInfo
            End If
        End If
        
        'adding textboxess
        'but we don't want to add header information
        'All headers contain an equals sign. We only add a textbox to play_shapes if it doesn't have this symbol
        'Any textbox that a user want to appear in a slide are not allowed to have an equals sign
        If shape.Type = msoTextBox Then
            If InStr(shape.TextFrame.TextRange.Text, "=") = 0 Then
                'Defensive players have one letter
                'Text has more than one letter
                If (Len(shape.TextFrame.TextRange.Text) = 1 And Not params.IgnoreDefenders) Or _
                    (Len(shape.TextFrame.TextRange.Text) > 1 And Not params.IgnoreText) Then
                    Set shapeInfo = New Dictionary
                    shapeInfo.add "Type", "TextBox"
                    shapeInfo.add "Label", shape.TextFrame.TextRange.Text
                    shapeInfo.add "Label Size", shape.TextFrame.TextRange.Font.Size
                    If params.Flip Then
                        If shape.left < centersShape.left Then
                            shapeInfo.add "Left", centersShape.left + centersShape.width + (centersShape.left - (shape.left + shape.width))
                        Else
                            shapeInfo.add "Left", centersShape.left - (shape.left - (centersShape.left + centersShape.width)) - shape.width
                        End If
                    Else
                        shapeInfo.add "Left", shape.left
                    End If
                    shapeInfo.add "Top", shape.top
                    shapeInfo.add "Width", shape.width
                    shapeInfo.add "Height", shape.height
                    shapeInfo.add "Text Color", shape.TextFrame.TextRange.Font.Color.RGB
                    copyShapes.add shapeInfo
                End If
            End If
        End If
        
        If shape.Type = msoLine Then
            If (shape.Line.DashStyle = msoLineSolid And Not params.IgnoreSolidLines) Or _
              (shape.Line.DashStyle <> msoLineSolid And Not params.IgnoreDashedLines) Then
                Set shapeInfo = New Dictionary
                
                'calculate coordinates of beginx, endx, beginy, endy
                Dim top As Single
                Dim left As Single
                Dim width As Single
                Dim height As Single
                top = shape.top
                left = shape.left
                width = shape.width
                height = shape.height
                            
                Dim beginX As Single
                Dim beginY As Single
                Dim endX As Single
                Dim endY As Single
                Dim lineOrientation As String
                lineOrientation = orientationOfLine(shape)
                
                If StrComp(lineOrientation, "DownRight", vbTextCompare) = 0 Then
                    beginX = left
                    beginY = top
                    endX = left + width
                    endY = top + height
                ElseIf StrComp(lineOrientation, "UpRight", vbTextCompare) = 0 Then
                    beginX = left
                    beginY = top + height
                    endX = left + width
                    endY = top
                ElseIf StrComp(lineOrientation, "DownLeft", vbTextCompare) = 0 Then
                    beginX = left + width
                    beginY = top
                    endX = left
                    endY = top + height
                Else 'StrComp(lineOrientation, "UpLeft", vbTextCompare) = 0 Then
                    beginX = left + width
                    beginY = top + height
                    endX = left
                    endY = top
                End If
                
                If params.Flip Then
                    Dim centersCenter As Single
                    centersCenter = centersShape.left + (centersShape.width / 2!)
                    If beginX < centersCenter Then
                        beginX = (centersCenter - beginX) + centersCenter
                    Else
                        beginX = centersCenter - (beginX - centersCenter)
                    End If
                    
                    If endX < centersCenter Then
                        endX = (centersCenter - endX) + centersCenter
                    Else
                        endX = centersCenter - (endX - centersCenter)
                    End If
                End If
                
                shapeInfo.add "Type", "Line"
                shapeInfo.add "BeginX", beginX
                shapeInfo.add "BeginY", beginY
                shapeInfo.add "EndX", endX
                shapeInfo.add "EndY", endY
                shapeInfo.add "Fore Color", shape.Line.ForeColor.RGB
                shapeInfo.add "Weight", shape.Line.Weight
                shapeInfo.add "Dash Style", shape.Line.DashStyle
                shapeInfo.add "Begin Arrow Style", shape.Line.BeginArrowheadStyle
                shapeInfo.add "End Arrow Style", shape.Line.EndArrowheadStyle
                copyShapes.add shapeInfo
            End If
        End If
    Next shape
    jsonObject.add "Shapes", copyShapes
    
    convertSlideToJson = JsonConverter.ConvertToJson(jsonObject, 2)
End Function



Public Sub addJsonShapesToSlide(ByRef slide As PowerPoint.slide, ByRef jsonString As String, Optional ByVal coreInjectionString As String = "12345")
    
    Dim jsonObject As Dictionary
    Set jsonObject = JsonConverter.ParseJson(jsonString)
    
    Dim shapeInfo As Dictionary
    For Each shapeInfo In jsonObject("Shapes")
        Dim newShape As PowerPoint.shape
        If shapeInfo("Type") = "Line" Then
            Set newShape = slide.shapes.AddLine(shapeInfo("BeginX"), shapeInfo("BeginY"), shapeInfo("EndX"), shapeInfo("EndY"))
            newShape.Line.ForeColor.RGB = shapeInfo("Fore Color")
            newShape.Line.Weight = shapeInfo("Weight")
            newShape.Line.DashStyle = shapeInfo("Dash Style")
            newShape.Line.BeginArrowheadStyle = shapeInfo("Begin Arrow Style")
            newShape.Line.EndArrowheadStyle = shapeInfo("End Arrow Style")
            
        End If
        If shapeInfo("Type") = "Oval" Then
            Set newShape = slide.shapes.AddShape(msoShapeOval, shapeInfo("Left"), shapeInfo("Top"), shapeInfo("Width"), shapeInfo("Height"))
            newShape.TextFrame.TextRange = performCoreInjection(shapeInfo("Label"), coreInjectionString)
            newShape.TextFrame.TextRange.Font.Size = shapeInfo("Label Size")
            newShape.Fill.ForeColor.RGB = shapeInfo("Fill Color")
            newShape.TextFrame.TextRange.Font.Color.RGB = shapeInfo("Text Color")
            newShape.Line.ForeColor.RGB = shapeInfo("Outline Color")
        End If
        If shapeInfo("Type") = "Rectangle" Then
            Set newShape = slide.shapes.AddShape(msoShapeRectangle, shapeInfo("Left"), shapeInfo("Top"), shapeInfo("Width"), shapeInfo("Height"))
            newShape.TextFrame.TextRange = performCoreInjection(shapeInfo("Label"), coreInjectionString)
            newShape.TextFrame.TextRange.Font.Size = shapeInfo("Label Size")
            newShape.Fill.ForeColor.RGB = shapeInfo("Fill Color")
            newShape.TextFrame.TextRange.Font.Color.RGB = shapeInfo("Text Color")
            newShape.Line.ForeColor.RGB = shapeInfo("Outline Color")
        End If
        If shapeInfo("Type") = "TextBox" Then
            Set newShape = slide.shapes.AddTextbox(msoTextOrientationHorizontal, shapeInfo("Left"), shapeInfo("Top"), shapeInfo("Width"), shapeInfo("Height"))
            newShape.TextFrame.TextRange = shapeInfo("Label")
            newShape.TextFrame.TextRange.Font.Size = shapeInfo("Label Size")
            newShape.TextFrame.TextRange.Font.Color.RGB = shapeInfo("Text Color")
        End If
        
    Next shapeInfo
    
End Sub


Public Function orientationOfLine(ByVal shape As PowerPoint.shape) As String
    Dim hFlip As Integer
    Dim vFlip As Integer

    hFlip = shape.HorizontalFlip
    vFlip = shape.VerticalFlip

    Select Case CStr(hFlip) & CStr(vFlip)
        Case "00"
            orientationOfLine = "DownRight"
        Case "0-1"
            orientationOfLine = "UpRight"
        Case "-10"
            orientationOfLine = "DownLeft"
        Case "-1-1"
            orientationOfLine = "UpLeft"
    End Select
End Function




