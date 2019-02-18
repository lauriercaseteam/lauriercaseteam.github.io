Sub DrawButtonBox(current_slide As Slide, dest_slide_idx As Integer, loc_x As Integer, loc_y As Integer, d As Integer)
    ' Inserts rectangle shape on current_slide in the top
    ' right corner with a link to dest_slide_idx.

    Dim oSh As Shape
    ' Shape Type, Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddShape(msoShapeRectangle, loc_x, loc_y, d, d)
    With oSh
        .Fill.ForeColor.RGB = RGB(215, 220, 230)
        .Line.Visible = False
    End With

    With oSh.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.SubAddress = dest_slide_idx
    End With

End Sub

Sub InsertTextBox(current_slide As Slide, text As String, loc_x As Integer, loc_y As Integer, h As Integer, w As Integer)
    ' Inserts rectangle textbox on current_slide at [loc_x, loc_y]
    ' position with centered string from text.

    Dim oSh As Shape
    ' Shape Type, Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, loc_x, loc_y, w, h)
    With oSh
        .Fill.ForeColor.RGB = RGB(0, 50, 99)
        .Line.Visible = False
    End With
    
    With oSh.TextFrame.TextRange
        .text = text
        .Paragraphs.ParagraphFormat.Alignment = 2
        With .Font
            .Size = 12
            .Name = "Arial"
            .Underline = msoTrue
            .Color = RGB(255, 255, 255)
        End With
    End With
    
End Sub

Sub InsertLinkTextbox(current_slide As Slide, dest_slide_idx As Integer, link_text As String, loc_x As Integer, loc_y As Integer, h As Integer, w As Integer)
    ' Inserts textbox shape on current_slide in the [loc_x, loc_y]
    ' position with a link to dest_slide_idx and text from link_text.

    Dim oSh As Shape
    ' Shape Type, Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, loc_x, loc_y, w, h)
    With oSh.TextFrame.TextRange
        .text = link_text
        With .Font
            .Size = 10
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
    ' Initialize all non-local variables.
    Dim app_slide_idx As Integer
    Dim extra_slide_idx As Integer
    Dim dest_slide As Integer
    Dim curr_x As Integer
    Dim curr_y As Integer
    Dim b_corner_x As Integer
    Dim b_corner_y As Integer
    Dim appendix_map_slide_num As Integer
    Dim ret_box_x As Integer
    Dim ret_box_y As Integer
    Dim link_box_h As Integer
    Dim link_box_w As Integer
    Dim return_box_dim As Integer

    Dim link_text As String

    Dim app_map_slide As Slide
    Dim extra_map_slide As Slide
    Dim current_slide As Slide

    Dim slide_h As Single
    Dim slide_w As Single

    ' Initialize key variables based on user input.
    num_slides = ActivePresentation.Slides.Count
    app_slide_idx = InputBox("What is the main appendices map cover slide number?") + 1
    extra_slide_idx = app_slide_idx + 1
    b1_div_idx = InputBox("What is the BUCKET 1 appendices cover slide number?")
    b2_div_idx = InputBox("What is the BUCKET 2 appendices cover slide number?")
    b3_div_idx = InputBox("What is the BUCKET 3 appendices cover slide number?")
    fin_div_idx = InputBox("What is the FINANCIAL appendices cover slide number?")
    extras_div_idx = InputBox("What is the EXTRA appendices cover slide number?")
    Set app_map_slide = ActivePresentation.Slides(app_slide_idx)
    Set extra_map_slide = ActivePresentation.Slides(extra_slide_idx)

    ' Check if we have any more slides after the appendix map slide.
    If num_slides = app_slide_idx Then
        MsgBox "No appendix slides found to link, make sure you put all your appendix slides after your appendix map slide i.e. slide #" & app_slide_idx & "."
        Exit Sub
    End If

    ' Get slide dimensions.
    With ActivePresentation.PageSetup
        slide_h = .SlideHeight
        slide_w = .SlideWidth
    End With
    
    ' Set height and width of textboxes for links
    link_box_h = 30
    link_box_w = 220
    
    ' Set dimensions for return boxes
    return_box_dim = 20
    
    ' Configurable parameters for locations of items in TOC.
    y_offset = 80
    x_offset = 30

    y_pos_incr = 40
    ' [(Width - borders on left and right) - (some room between columns * 3 buckets)] / 3 buckets
    x_pos_incr = ((slide_w - x_offset * 2) - (x_offset * 3)) / 3
    
    ' Set location of return button box. (20x20) is the hardcoded size of return button.
    ret_box_x = slide_w - return_box_dim
    ret_box_y = 0
    
    ' Configurable parameters for location of extra appendices link button.
    b_corner_x = slide_w - x_offset - link_box_w
    b_corner_y = slide_h - y_offset - link_box_h
    linked_slides = 0
    
    ' ############## Link Bucket 1 Appendices ##############
    curr_y = y_offset
    curr_x = x_offset
    link_text = "Bucket 1"
    InsertTextBox app_map_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
    
    For i = b1_div_idx + 1 To b2_div_idx - 1
        link_text = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.text
        dest_slide = i

        ' Get location for next textbox.
        curr_y = curr_y + y_pos_incr
        
        InsertLinkTextbox app_map_slide, dest_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
        linked_slides = linked_slides + 1
    Next i
    
    ' ############## Link Bucket 2 Appendices ##############
    curr_y = y_offset
    curr_x = curr_x + x_pos_incr + x_offset
    link_text = "Bucket 2"
    InsertTextBox app_map_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
    
    For i = b2_div_idx + 1 To b3_div_idx - 1
        link_text = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.text
        dest_slide = i
        curr_y = curr_y + y_pos_incr
        
        InsertLinkTextbox app_map_slide, dest_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
        linked_slides = linked_slides + 1
    Next i
    
    ' ############## Link Bucket 3 Appendices ##############
    curr_y = y_offset
    curr_x = curr_x + x_pos_incr + x_offset
    link_text = "Bucket 3"
    InsertTextBox app_map_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
    
    For i = b3_div_idx + 1 To fin_div_idx - 1
        link_text = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.text
        dest_slide = i
        curr_y = curr_y + y_pos_incr
        
        InsertLinkTextbox app_map_slide, dest_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
        linked_slides = linked_slides + 1
    Next i
    
    ' ############## Link Financial Appendices ##############
    curr_y = y_offset
    curr_x = x_offset
    link_text = "Financials"
    InsertTextBox extra_map_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
    
    For i = fin_div_idx + 1 To extras_div_idx - 1
        link_text = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.text
        dest_slide = i
        curr_y = curr_y + y_pos_incr
        
        InsertLinkTextbox extra_map_slide, dest_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
        linked_slides = linked_slides + 1
    Next i
    
    ' ############## Link Extra Appendices ##############
    curr_y = y_offset
    curr_x = curr_x + x_pos_incr * 2 + x_offset * 2
    link_text = "Extras"
    InsertTextBox extra_map_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
    
    For i = extras_div_idx + 1 To num_slides
        link_text = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.text
        dest_slide = i
        curr_y = curr_y + y_pos_incr
        
        InsertLinkTextbox extra_map_slide, dest_slide, link_text, curr_x, curr_y, link_box_h, link_box_w
        linked_slides = linked_slides + 1
    Next i

    ' Add return button to main appendices slide.
    For i = 1 To fin_div_idx - 1
        Set current_slide = ActivePresentation.Slides(i)
        DrawButtonBox current_slide, app_slide_idx, ret_box_x, ret_box_y, return_box_dim
    Next i
    
    ' Add return button to extra appendices slide.
    For i = fin_div_idx + 1 To num_slides
        Set current_slide = ActivePresentation.Slides(i)
        DrawButtonBox current_slide, extra_slide_idx, ret_box_x, ret_box_y, return_box_dim
    Next i
    
    ' Add link button to go from main appendices to extras
    link_text = "Extras >"
    InsertLinkTextbox app_map_slide, extra_slide_idx, link_text, b_corner_x, b_corner_y, link_box_h, link_box_w
    
    ' Add link button to go from extra appendices to main
    link_text = "< Main"
    InsertLinkTextbox extra_map_slide, app_slide_idx, link_text, b_corner_x, b_corner_y, link_box_h, link_box_w
    

    ' Print final message on number of linked slides.
    msg_end = ""
    If linked_slides > 1 And (num_slides - app_slide_idx) = linked_slides Then
        msg_end = " slides."
    ElseIf (num_slides - app_slide_idx) > linked_slides Then
        msg_end = " slides. Exited early."
    Else
        msg_end = " slide."
    End If

    MsgBox "Finished linking " & linked_slides & msg_end
End Sub
