"""
Creating powerpoint presentation containing texts and images
using python
"""

import win32com.client as win

# Creating RGB function
def RGB(red, green, blue):
    assert 0 <= red <=255    
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)

app = win.Dispatch("PowerPoint.Application")
ppt = app.Presentations.Add(True)

# Length and width of ppt slides
l = ppt.PageSetup.SlideHeight
b = ppt.PageSetup.SlideWidth

filename = 'dummy.ppt'
pictname = 'dummy.jpg'

pptfile = ppt.SaveAs(filename, 1)

# Adding slide and then a textbox into it
slide1 = ppt.Slides.AddSlide(1,12)
shape1 = slide1.Shapes.AddTextbox(1,20,20,20,20)
shape1.TextFrame.TextRange.Text = "Dummy Presentation"

# Fonts formatting in textbox
shape1.TextFrame.TextRange.Font.Size = 24
shape1.TextFrame.TextRange.Font.Bold = True
shape1.TextFrame.TextRange.Font.Name = "Palatino"
shape1.TextFrame.TextRange.Font.color.RGB = RGB(0,0,255)

shape1.Fill.ForeColor.RGB = RGB(128, 0, 0) 
shape1.Fill.BackColor.RGB = RGB(170, 170, 170)

# Slide to insert image
slide2 = ppt.Slides.AddSlide(2,12)
slide2.Shapes.AddPicture(FileName=pictName, LinkToFile=False,
            SaveWithDocument=True, Left=100, Top=100, Width=200, Height=200)

ppt.Save()