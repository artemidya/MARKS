Private Function RemoveCoincidentLineSegmentToLineSegment(TrimmingSegment As VGCore.segment, BreakSegment As VGCore.segment) As Long
    'Line-Line
    'x--------->  BreakSegment
    'x=========>  TrimmingSegment

    If (DeltaAnglBetweenSegments(BreakSegment, TrimmingSegment)) < ThresholdAngle Then
        Dim PT As ParametrT2X
        Dim dist1 As Distance2X
        Dim dist2 As Distance2X
        
        PT = ParametrT(BreakSegment, TrimmingSegment)
        dist1 = Distance(BreakSegment, TrimmingSegment)
        dist2 = Distance(TrimmingSegment, BreakSegment)
        Dim delNode As VGCore.Node
        Dim t As Double
         
        With PT
            If .t2 > .t1 Then
                '>>>>>>
                If dist1.Distance1 < Sensitivity Or dist2.Distance1 < Sensitivity Then
                    If dist1.Distance2 < Sensitivity Or dist2.Distance2 < Sensitivity Then
                        Select Case ((PT.t1 > 0 And PT.t1 < 1) * 2 + (PT.t2 > 0 And PT.t2 < 1))
                            Case 0 'false false
                                'x--------->  x====>, x====>   x------->
                                If PT.t1 <= 0 And PT.t2 >= 1 Then
                                    'x====x====>======>, xx===>>
                                    Set delNode = BreakSegment.EndNode
                                    BreakSegment.EndNode.BreakApart
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = -1
                                End If
                            Case -1 'false true
                                'x==x=>------->
                                Set delNode = BreakSegment.StartNode
                                BreakSegment.BreakApartAt PT.t2
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -2 'true false
                                'x------x==>==>
                                Set delNode = BreakSegment.EndNode
                                BreakSegment.BreakApartAt PT.t1
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -3 'true true
                                'x--x====>->
                                t = (PT.t2 - PT.t1) / (1 - PT.t1)
                                If t > 0 And t < 1 Then
                                    Set delNode = BreakSegment.BreakApartAt(PT.t1)
                                    delNode.NextSegment.BreakApartAt t
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = 1
                                End If
                        End Select
                    End If
                End If
            Else
                '>>><<<
                If dist1.Distance1 < Sensitivity Or dist2.Distance2 < Sensitivity Then
                    If dist1.Distance2 < Sensitivity Or dist2.Distance1 < Sensitivity Then
                        Select Case (PT.t2 > 0 And PT.t2 < 1) * 2 + (PT.t1 > 0 And PT.t1 < 1)
                            Case 0 'false false
                                'x--------->  <====x, x====>   <-------x
                                If PT.t2 <= 0 And PT.t1 >= 1 Then
                                    '<====x====>======x, <x===>x
                                    Set delNode = BreakSegment.EndNode
                                    BreakSegment.EndNode.BreakApart
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = -1
                                End If
                            Case -1 'false true
                                '<==x===x----->
                                Set delNode = BreakSegment.StartNode
                                BreakSegment.BreakApartAt PT.t1
                                BreakSegment.EndNode.Delete
'                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -2 'true false
                                'x------<==>===x
                                Set delNode = BreakSegment.EndNode
                                BreakSegment.BreakApartAt PT.t2
                                delNode.Next.Delete
                                'RemoveCoincidentLineSegmentToLineSegment = 0
                            Case -3 'true true
                                'x--<====x->
                                t = (PT.t1 - PT.t2) / (1 - PT.t2)
                                If t > 0 And t < 1 Then
                                    Set delNode = BreakSegment.BreakApartAt(PT.t2)
                                    delNode.NextSegment.BreakApartAt t
                                    delNode.Delete
                                    RemoveCoincidentLineSegmentToLineSegment = 1
                                End If
                        End Select
                    End If
                End If
            End If
        End With
    End If
End Function
