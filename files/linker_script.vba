Sub DrawButtonBox(current_slide As Slide, dest_slide_idx As Integer)
    Dim oSh As Shape
    ' Left offset, Top offset, Height, Width
    Set oSh = current_slide.Shapes.AddShape(msoShapeRectangle, 940, 0, 20, 20)
    With oSh
        .Fill.ForeColor.RGB = RGB(214, 220, 229)
        .Line.Visible = False
    End With
    
    With oSh.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.SubAddress = dest_slide_idx
    End With
    
End Sub

Sub InsertLinkTextbox(current_slide As Slide, dest_slide_idx As Integer, link_text As String, loc_x As Integer, loc_y As Integer)
    Dim oSh As Shape
    ' Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, loc_x, loc_y, 200, 30)
    With oSh.TextFrame.TextRange
    .text = link_text
    With .Font
            .Size = 14
            .Name = "Arial"
            .Underline = msoTrue
        End With
    End With
    
    With oSh.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.SubAddress = dest_slide_idx
    End With
End Sub

Sub linker()
    ' Link all text boxes on the slideasdfasdf
    Dim app_slide_idx As Integer
    Dim dest_slide As Integer
    Dim curr_x As Integer
    Dim curr_y As Integer
    Dim link_text As String
    Dim appendix_slide As Slide
    Dim appendix_map_slide_num As Integer
    Dim current_slide As Slide
    
    num_slides = ActivePresentation.Slides.Count
    app_slide_idx = InputBox("What is the appendix slide number?")
    Set appendix_slide = ActivePresentation.Slides(app_slide_idx)
    
    y_start_pos = 75
    x_start_pos = 30
    y_cutoff_pos = 400
    
    y_pos_incr = 30
    x_pos_incr = 300
    
    curr_y = y_start_pos - y_pos_incr
    curr_x = x_start_pos

    ' Go over all slides after the appendix map slide
    For i = app_slide_idx + 1 To num_slides
    
        ' For every slide, ask what text box text does the user want
        box_prompt = "What is the link title for slide " & i & " ?" & " If you want to skip to the next slide, leave input box blank and press 'ok'"
        link_text = InputBox(box_prompt)
        dest_slide = i
        
        ' Get location for next textbox
        If curr_y + y_pos_incr >= y_cutoff_pos Then
            curr_x = curr_x + x_pos_incr
            curr_y = y_start_pos
        Else
            curr_y = curr_y + y_pos_incr
        End If
        
        ' Parse user input
        If StrPtr(link_text) = 0 Then
            Exit For
        ElseIf link_text = "" Then
            ' If we provide no input, treat it as skipping on to the next slide
        Else
            InsertLinkTextbox appendix_slide, dest_slide, link_text, curr_x, curr_y
        End If
        
    Next i
    
    ' Add return button to all slides
    For i = 1 To num_slides
        Set current_slide = ActivePresentation.Slides(i)
        DrawButtonBox current_slide, app_slide_idx
    Next i
End Sub

