' 第01课《数字的奥秘》PPT动画宏代码
' 卡通风格动画效果
' 使用方法：在PowerPoint中按Alt+F11打开VBA编辑器，导入此文件并运行

Public Sub AddAllAnimations()
    ' 为所有幻灯片添加动画效果
    Dim slide As slide
    Dim shape As shape
    Dim delay As Single
    
    For Each slide In ActivePresentation.Slides
        delay = 0
        
        Select Case slide.SlideIndex
            Case 1 ' 封面页
                AddCoverAnimations slide
            Case 2 ' 课程目标页
                AddObjectiveAnimations slide
            Case 3, 6, 7 ' 例题页
                AddExampleAnimations slide
            Case 4, 5 ' 概念页
                AddConceptAnimations slide
            Case 8 ' 练习页
                AddPracticeAnimations slide
            Case 9 ' 总结页
                AddSummaryAnimations slide
            Case 10 ' 作业页
                AddHomeworkAnimations slide
            Case 11 ' 结束页
                AddEndingAnimations slide
        End Select
    Next slide
    
    MsgBox "动画添加完成！🎉", vbInformation, "夏老师数学课堂"
End Sub

Private Sub AddCoverAnimations(slide As slide)
    ' 封面页动画 - 弹跳+淡入效果
    Dim shape As shape
    
    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                If shape.TextFrame.TextRange.Font.Size >= 36 Then
                    ' 大标题弹跳效果
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectBounce
                        .Timing.TriggerDelayTime = 0.2
                    End With
                Else
                    ' 副标题淡入效果
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectFade
                        .Timing.TriggerDelayTime = 0.5
                    End With
                End If
            End If
        End If
    Next shape
End Sub

Private Sub AddObjectiveAnimations(slide As slide)
    ' 课程目标页 - 卡片飞入效果
    Dim shape As shape
    Dim delay As Single
    delay = 0
    
    For Each shape In slide.Shapes
        If shape.Type = msoPlaceholder Then
            With shape.AnimationSettings
                .EntryEffect = ppEffectFlyFromBottom
                .Timing.TriggerDelayTime = delay
            End With
            delay = delay + 0.3
        End If
    Next shape
End Sub

Private Sub AddExampleAnimations(slide As slide)
    ' 例题页 - 逐步显示
    Dim shape As shape
    Dim delay As Single
    delay = 0
    
    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            With shape.AnimationSettings
                .EntryEffect = ppEffectPeekFromLeft
                .Timing.TriggerDelayTime = delay
            End With
            delay = delay + 0.4
        End If
    Next shape
End Sub

Private Sub AddConceptAnimations(slide As slide)
    ' 概念页 - 缩放效果
    Dim shape As shape
    Dim delay As Single
    delay = 0
    
    For Each shape In slide.Shapes
        With shape.AnimationSettings
            .EntryEffect = ppEffectZoomCenter
            .Timing.TriggerDelayTime = delay
        End With
        delay = delay + 0.25
    Next shape
End Sub

Private Sub AddPracticeAnimations(slide As slide)
    ' 练习页 - 旋转飞入
    Dim shape As shape
    Dim delay As Single
    delay = 0
    
    For Each shape In slide.Shapes
        If shape.Type = msoPlaceholder Then
            With shape.AnimationSettings
                .EntryEffect = ppEffectFlyFromRight
                .Timing.TriggerDelayTime = delay
            End With
            delay = delay + 0.35
        End If
    Next shape
End Sub

Private Sub AddSummaryAnimations(slide As slide)
    ' 总结页 - 强调效果
    Dim shape As shape
    
    For Each shape In slide.Shapes
        With shape.AnimationSettings
            .EntryEffect = ppEffectGrowShrink
            .Timing.TriggerDelayTime = 0.3
        End With
    Next shape
End Sub

Private Sub AddHomeworkAnimations(slide As slide)
    ' 作业页 - 层叠效果
    Dim shape As shape
    Dim delay As Single
    delay = 0
    
    For Each shape In slide.Shapes
        With shape.AnimationSettings
            .EntryEffect = ppEffectCrawlFromUp
            .Timing.TriggerDelayTime = delay
        End With
        delay = delay + 0.4
    Next shape
End Sub

Private Sub AddEndingAnimations(slide As slide)
    ' 结束页 - 华丽效果
    Dim shape As shape
    
    For Each shape In slide.Shapes
        If shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                If shape.TextFrame.TextRange.Font.Size >= 32 Then
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectSpin
                        .Timing.TriggerDelayTime = 0.2
                    End With
                Else
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectFade
                        .Timing.TriggerDelayTime = 0.6
                    End With
                End If
            End If
        End If
    Next shape
End Sub

Public Sub ClearAllAnimations()
    ' 清除所有动画，防止重复添加
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            shape.AnimationSettings.EntryEffect = ppEffectNone
        Next shape
    Next slide
    
    MsgBox "所有动画已清除！", vbInformation, "夏老师数学课堂"
End Sub

Public Sub AddCartoonTitleBounce()
    ' 为所有标题添加卡通弹跳动画
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    If shape.TextFrame.TextRange.Font.Size >= 36 Then
                        With shape.AnimationSettings
                            .EntryEffect = ppEffectBounce
                            .Timing.TriggerDelayTime = 0.3
                        End With
                    End If
                End If
            End If
        Next shape
    Next slide
End Sub

Public Sub AddCuteCardFlyIn()
    ' 为卡片添加可爱飞入动画
    Dim slide As slide
    Dim shape As shape
    Dim delay As Single
    
    For Each slide In ActivePresentation.Slides
        delay = 0
        For Each shape In slide.Shapes
            If shape.Type = msoAutoShape Then
                If shape.AutoShapeType = msoShapeRoundedRectangle Then
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectFlyFromBottom
                        .Timing.TriggerDelayTime = delay
                    End With
                    delay = delay + 0.2
                End If
            End If
        Next shape
    Next slide
End Sub

Public Sub AddNumberSequenceAnimation()
    ' 为数列添加逐个显示动画
    Dim slide As slide
    Dim shape As shape
    Dim delay As Single
    
    For Each slide In ActivePresentation.Slides
        delay = 0
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If InStr(shape.TextFrame.TextRange.Text, ",") > 0 Then
                    ' 可能是数列
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectAppear
                        .Timing.TriggerDelayTime = delay
                    End With
                    delay = delay + 0.5
                End If
            End If
        Next shape
    Next slide
End Sub

Public Sub AddEmojiPulse()
    ' 为emoji添加脉动效果
    Dim slide As slide
    Dim shape As shape
    
    For Each slide In ActivePresentation.Slides
        For Each shape In slide.Shapes
            If shape.HasTextFrame Then
                If ContainsEmoji(shape.TextFrame.TextRange.Text) Then
                    With shape.AnimationSettings
                        .EntryEffect = ppEffectGrowShrink
                        .Timing.TriggerDelayTime = 0.3
                    End With
                End If
            End If
        Next shape
    Next slide
End Sub

Private Function ContainsEmoji(text As String) As Boolean
    ' 检查文本是否包含emoji
    Dim emojiChars As Variant
    Dim i As Integer
    
    emojiChars = Array("🔢", "🎯", "🎮", "📚", "💡", "📝", "🌟", "🎁", "👋", _
                       "🐸", "🔐", "🧩", "🐰", "🎡", "⭐", "🌸", "🍀", "🔴", "🔵", "🟡")
    
    For i = LBound(emojiChars) To UBound(emojiChars)
        If InStr(text, emojiChars(i)) > 0 Then
            ContainsEmoji = True
            Exit Function
        End If
    Next i
    
    ContainsEmoji = False
End Function

Public Sub PreviewAllEffects()
    ' 预览所有动画效果
    Dim slide As slide
    
    For Each slide In ActivePresentation.Slides
        slide.SlideShowTransition.EntryEffect = ppEffectRandom
        slide.SlideShowTransition.Duration = 1
    Next slide
    
    MsgBox "所有幻灯片切换效果已设置为随机！开始放映查看效果吧！", vbInformation, "夏老师数学课堂"
End Sub
