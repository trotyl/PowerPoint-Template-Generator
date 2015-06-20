Sub Generate()
    Dim slideCount As Integer, slideTemplate As Variant
    Let slideCount = ActivePresentation.Slides.Count
    Set slideTemplate = ActivePresentation.Slides(slideCount)
    Dim str As String
    Let str = InputBox("请输入Excel路径", "替换内容")
    Open str For Input As #1
    Do While Not EOF(1)
        Dim newSlide As Variant
        Set newSlide = slideTemplate.Duplicate
        Let slideCount = slideCount + 1
        newSlide.MoveTo slideCount
        ActiveWindow.View.GotoSlide newSlide.SlideIndex
        Dim strline As String
        Line Input #1, strline
        Dim contentArr As Variant
        Let contentArr = Split(strline, Chr(9))
        Dim content As Variant, index As Integer
        Let index = 0
        For Each content In contentArr
            Dim control As Variant
            For Each control In newSlide.Shapes
                If control.HasTextFrame Then
                    Dim tmp As String
                    Let tmp = control.TextFrame.TextRange.Text
                    control.TextFrame.TextRange.Text = Replace(tmp, "{" & index & "}", content)
                End If
            Next
            Let index = index + 1
        Next
    
    Loop
    Close (1)
End Sub



