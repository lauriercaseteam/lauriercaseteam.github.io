Sub DrawButtonBox(current_slide As Slide, dest_slide_idx As Integer, loc_x As Integer, loc_y As Integer)
    ' Inserts rectangle shape on current_slide in the top
    ' right corner with a link to dest_slide_idx.

    Dim oSh As Shape
    ' Shape Type, Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddShape(msoShapeRectangle, loc_x, loc_y, 20, 20)
    With oSh
        .Fill.ForeColor.RGB = RGB(215, 220, 230)
        .Line.Visible = False
    End With

    With oSh.ActionSettings(ppMouseClick)
        .Action = ppActionHyperlink
        .Hyperlink.SubAddress = dest_slide_idx
    End With

End Sub

Sub InsertLinkTextbox(current_slide As Slide, dest_slide_idx As Integer, link_text As String, loc_x As Integer, loc_y As Integer)
    ' Inserts textbox shape on current_slide in the [loc_x, loc_y]
    ' position with a link to dest_slide_idx and text from link_text.

    Dim oSh As Shape
    ' Shape Type, Left offset, Top offset, Width, Height
    Set oSh = current_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, loc_x, loc_y, 200, 30)
    With oSh.TextFrame.TextRange
    .Text = link_text
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
    ' Initialize all non-local variables.
    Dim app_slide_idx As Integer
    Dim dest_slide As Integer
    Dim curr_x As Integer
    Dim curr_y As Integer
    Dim appendix_map_slide_num As Integer
    Dim ret_box_x As Integer
    Dim ret_box_y As Integer

    Dim link_text As String

    Dim appendix_slide As Slide
    Dim current_slide As Slide

    Dim slide_h As Single
    Dim slide_w As Single

    ' Initialize key variables based on user input.
    num_slides = ActivePresentation.Slides.Count
    app_slide_idx = InputBox("What is the appendix slide number?")
    Set appendix_slide = ActivePresentation.Slides(app_slide_idx)

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

    ' Set location of return button box. (20x20) is the hardcoded size of return button.
    ret_box_x = slide_w - 20
    ret_box_y = 0

    ' Configurable parameters for locations of items in TOC.
    y_start_pos = 75
    x_start_pos = 30
    y_cutoff_pos = 400

    y_pos_incr = 30
    x_pos_incr = 300

    curr_y = y_start_pos - y_pos_incr
    curr_x = x_start_pos

    linked_slides = 0

    ' Go over all slides after the appendix map slide.
    For i = app_slide_idx + 1 To num_slides
        slide_title = ActivePresentation.Slides(i).Shapes.Title.TextFrame.TextRange.Text

        ' For every slide, ask what text box text does the user want.
        skip_msg = vbCrLf & vbCrLf & "If you want to skip to the next slide, leave input box blank and press 'ok'"
        box_prompt = "What is the link title for slide " & i & ":  " & slide_title & "?" & skip_msg
        link_text = InputBox(box_prompt)
        dest_slide = i

        ' Get location for next textbox.
        If curr_y + y_pos_incr >= y_cutoff_pos Then
            curr_x = curr_x + x_pos_incr
            curr_y = y_start_pos
        Else
            curr_y = curr_y + y_pos_incr
        End If

        ' Parse user input.
        If StrPtr(link_text) = 0 Then
            Exit For
        ElseIf link_text = "" Then
            ' If we provide no input, treat it as skipping on to the next slide.
        Else
            InsertLinkTextbox appendix_slide, dest_slide, link_text, curr_x, curr_y
            linked_slides = linked_slides + 1
        End If

    Next i

    ' Add return button to all slides.
    For i = 1 To num_slides
        Set current_slide = ActivePresentation.Slides(i)
        DrawButtonBox current_slide, app_slide_idx, ret_box_x, ret_box_y
    Next i

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
