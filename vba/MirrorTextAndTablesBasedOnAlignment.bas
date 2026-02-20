Attribute VB_Name = "MirrorTextAndTablesBasedOnAlignment"
' RTL mirror: ungroup, mirror positions (except title), toggle text/table direction.
' Use from Python via AppleScript: run VB macro "MirrorTextAndTablesBasedOnAlignment"

Sub MirrorTextAndTablesBasedOnAlignment()
    Dim sld As slide
    Dim shp As Shape
    Dim slideWidth As Single
    Dim groupsExist As Boolean: groupsExist = True
    Dim firstTextBoxFound As Boolean

    Set sld = ActiveWindow.View.slide
    slideWidth = sld.Master.Width  ' Get the slide width from the master layout

    Do While (groupsExist = True)
      groupsExist = False
      For Each shp In sld.Shapes
          If shp.Type = msoGroup Then
              shp.Ungroup
              intCount = intCount + 1
              groupsExist = True
          End If
      Next shp
    Loop

    slide_title = sld.Shapes.Title.TextFrame.TextRange.Text

    For Each shp In sld.Shapes

        If shp.HasTextFrame Then
            If slide_title = shp.TextFrame.TextRange.Text Then

                If shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                    shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                ElseIf shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft Then
                    shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight
                End If

            ElseIf shp.HasTextFrame Then
                shp.LockAspectRatio = msoTrue  ' Lock aspect ratio to maintain proportions
                shp.Left = slideWidth - shp.Left - shp.Width  ' Mirror the position along the vertical axis

                If shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                    shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                ElseIf shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft Then
                    shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight
                End If

            Else
                shp.LockAspectRatio = msoTrue  ' Lock aspect ratio to maintain proportions
                shp.Left = slideWidth - shp.Left - shp.Width  ' Mirror the position along the vertical axis

            End If

        ElseIf shp.Type = msoTable Then
            shp.LockAspectRatio = msoTrue  ' Lock aspect ratio to maintain proportions
            shp.Left = slideWidth - shp.Left - shp.Width  ' Mirror the position along the vertical axis

            If shp.Table.TableDirection = ppDirectionLeftToRight Then
                shp.Table.TableDirection = ppDirectionRightToLeft

            ElseIf shp.Table.TableDirection = ppDirectionRightToLeft Then
                shp.Table.TableDirection = ppDirectionLeftToRight
            End If

            Dim row As row
            Dim cell As cell

            ' Change text direction of all cells in the table
            For Each row In shp.Table.Rows
                For Each cell In row.Cells
                    If cell.Shape.HasTextFrame Then
                        If cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                            cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
                        ElseIf cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft Then
                            cell.Shape.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight
                        End If
                    End If
                Next cell
            Next row

        ElseIf shp.HasTextFrame Then
            shp.LockAspectRatio = msoTrue  ' Lock aspect ratio to maintain proportions
            shp.Left = slideWidth - shp.Left - shp.Width  ' Mirror the position along the vertical axis

            If shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight Then
                shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft
            ElseIf shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionRightToLeft Then
                shp.TextFrame.TextRange.ParagraphFormat.TextDirection = msoTextDirectionLeftToRight
            End If

        Else
            shp.LockAspectRatio = msoTrue  ' Lock aspect ratio to maintain proportions
            shp.Left = slideWidth - shp.Left - shp.Width  ' Mirror the position along the vertical axis

        End If

    Next shp
End Sub
