---
layout: default
title: PowerPoint
nav_order: 1
has_children: false
---

Det er jo oplagt at oprette PowerPoint præsentationer baseret på data fra et Excel ark.

## Åben PowerPoint

```vbnet
Sub AabenPowerPoint()
    ' Variable
    Dim appPowerPoint As Object

    ' Åben PowerPoint
    Set appPowerPoint = CreateObject("PowerPoint.Application")
    
    ' Handlinger i PowerPoint
    With appPowerPoint
        .Visible = True
        .Presentations.Add
        .ActivePresentation.Slides.Add 1, 1
    End With

End Sub
```

[Slide.Layout property (PowerPoint)](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.slide.layout)

## Åben eksisterende præsentation
Hvis du har en PowerPoint præsentation du vil åbne, kan du gøre det på denne måde.  
Bemærk at du skal angive stien og filnavnet i variablen *PowerPointPress*

```vbnet
Sub AabenPowerPointPresentation()
    ' Variable
    Dim appPowerPoint As Object
    Dim PowerPointPress As String

    ' Eksisterende PowerPoint præsentation
    PowerPointPress = "C:\Users\Tue Hellstern\Documents\Salgsdata.pptx"
    
    ' Åben PowerPoint
    Set appPowerPoint = CreateObject("PowerPoint.Application")
    
    ' Handlinger i PowerPoint
    With appPowerPoint
        .Visible = True
        .Presentations.Open (PowerPointPress)
    End With

End Sub
```