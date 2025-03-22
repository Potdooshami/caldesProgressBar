' Log
'2024.11.15: 일단 글로우 추가.

Sub caldesProgressBar()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim atom As Shape
    Dim wf As Shape
    Dim atomR As Single
    Dim pts() As Single
    Dim totalSlides As Long
    Dim posAtomX As Single
    Dim posAtomY As Single
    Dim amps(1 To 3) As Single
    Dim indAtoms As Integer
    
    ' 여기부터 손으로 건드릴 부분
    atomR = 5
    atom2R = 2 * atomR
    ibsRed = RGB(175, 39, 47)
    ibsGray = RGB(99, 102, 106)
    Set ppt = ActivePresentation
    sw = ppt.PageSetup.SlideWidth
    sh = ppt.PageSetup.SlideHeight
    totalSlides = ppt.Slides.Count
    xOff = (sw - atom2R * totalSlides) / 2
    yOff = sh - atom2R - 3
    nAtom = totalSlides
    nPoint = 2 * nAtom ' + 1
    nCtrl = 3 * nPoint + 1
    ampRed = atomR / 2
    ampSoliton = atom2R
    ampGray = atomR
    squareness = atomR / 2
    amps(3) = ampRed
    amps(2) = ampSoliton
    amps(1) = ampGray
    ReDim pts(1 To nCtrl, 1 To 2) As Single
    For indSlide = 1 To totalSlides
        For indAtoms = 1 To totalSlides
            posAtomX = (indAtoms - 1) * atom2R + xOff
            posAtomY = yOff
            Set atom = ActivePresentation.Slides(indSlide).Shapes.AddShape(msoShapeOval, posAtomX, posAtomY, atom2R, atom2R)
            If indAtoms > indSlide Then
                atom.Fill.ForeColor.RGB = ibsRed
                atom.Line.ForeColor.RGB = RGB(100, 0, 0)
            Else
                atom.Fill.ForeColor.RGB = ibsGray
                atom.Line.ForeColor.RGB = RGB(20, 20, 50)
            End If
            
            ' atom.Line.Visible = msoFalse
        Next indAtoms
        For indCntrl = 1 To nCtrl
            indPoint = (indCntrl \ 3) + 1
            indAtom = indPoint \ 2
            pts(indCntrl, 1) = (indPoint - 1) * atomR + squareness * ((indCntrl Mod 3) - 1) + xOff
            pts(indCntrl, 2) = -((indPoint + 1) Mod 2) * amps(Sgn(indSlide - indAtom) + 2) + yOff
        Next indCntrl
        Set wf = ActivePresentation.Slides(indSlide).Shapes.AddCurve(pts)
        wf.Line.ForeColor.RGB = RGB(255, 0, 0)
        wf.Line.Transparency = 0.99
        wf.Line.Weight = 1
        wf.Glow.Color.RGB = RGB(255, 0, 0)
        wf.Glow.Radius = 2.5
        wf.Glow.Transparency = 0.5
        
        
        
        
        
        
        
    Next indSlide
    
    
    ' indSlide = 10
    
    
    
    
    
End Sub

