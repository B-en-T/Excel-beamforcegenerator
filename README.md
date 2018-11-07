


# Excel-beamforcegenerator
Generate Graphs of forces in Excel and calculate Stress in fixed and rolling support beam situations


![alt text](https://i.imgur.com/MPletqB.jpg)



# Word in advance

 This class to generate force and tension graphs was made for a school assignment
 The code in question has been written quite fast and is in no way optimized
 I took shortcuts and also i am to unexpirienced to write quality VB code.

 be aware for the existence of bugs and possible slowdowns/ delays / crashes by non optimized code
 I mostly tested the fixed situation option and it worked for me, hope you can use this or put it to good use

 a well if you hate it... at least it was free

 thanks to all tutorials and help topics about VB/Excel that i needed to make this

 Best wishes, Tony


 # Setup / Usage

 create a class module in exel and copy atleast all of the non quoted code to it
 name / rename this class to modelerenclass
  
 
add this quoted code or a similar code to your excel vb excel project under modules -> modulename (probably Module1)

-----------------------------------------------------

```
Function createmodel(classname As String, berekentype As String, sheetname As String, forcepos As String, momentpos As String, supportpos As String, overigpos As String, calcpos As String, efpos As String)

   Dim mClass As New modelerenclass
   Dim lastactivecell As String
   Dim curSheetname As String

   'sla huidige sheet en cel op
   lastactivecell = ActiveCell.Address
   curSheetname = ActiveSheet.name
   Application.EnableEvents = False


    mClass.rundefault classname, berekentype, sheetname, forcepos, momentpos, supportpos, overigpos, calcpos, efpos



    'herselecteer de sheet en cel
    Application.EnableEvents = True
   Range(lastactivecell).Select
    Worksheets(curSheetname).Activate


End Function
```



 you can create all tables on a new sheet with the following line of code

```
createmodel "Uniquesheetname", "type of calculation", "Sheetname to add the tables",  "A3", "A6", "A9", "A13", "I7", "I3"
```

 example 

```
 createmodel "mClasstest", "fixed", "Sheetname", "A3", "A6", "A9", "A13", "I7", "I3"
```

calling this line should generate empty tables on the sheetname
calling this line again should update the graphs and tables by checking for filled in values inside the tables


options explanation

type of calculation options: "fixed" (fixed point on the left) or "punt" (hinge + rolling point calculation)
A third option was under development called "AxBx" (two fixed points and usage of degrees for FA,FB) this one doesnt work properly,
(i don't /didnt needed it anyway)'
the positions of generated table positions are changeable by changing the cornercell names -> "A3", "A6", "A9", "A13", "I7", "I3"



# on table change Auto run code example

usage and Steps

add this code or a similar code to your excel vb excel project under modules -> modulename (probably Module1)
to add more sheets just reuse / copy paste the code inside the sub


```
Sub runmodels(sheetedit As String)
  
   If sheetedit = "Sheetname" Or sheetedit = "all" Then
    
    createmodel "mClass", "fixed", "Sheetname", "A3", "A6", "A9", "A13", "I7", "I3"

   End If
  
  
End Sub
```


add this code or a similar code to your excel vb excel project under  microsoft excel objects --> ThisWorkbook


```
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    'Automatically updates any table that has "UpdatedBy" and "UpdatedOn" columns
    
    Dim c                               As Range
    Dim lo                              As ListObject

    Application.EnableEvents = False
    
    For Each c In Target.Cells
        'Debug.Print Now & "  " & Sh.Name, c.Address, """" & c.Value & """"
        For Each lo In Sh.ListObjects
        On Error GoTo Exit_Workbook_SheetChange
            If Not Intersect(c, lo.DataBodyRange) Is Nothing Then
                
                
                   Call runmodels(Sh.name)
                   
                   
                On Error GoTo 0
                Exit For 'lo
            End If
        Next lo
    Next c
    
Exit_Workbook_SheetChange:
    Set lo = Nothing
    Set c = Nothing
                
    Application.EnableEvents = True
    
End Sub
```


create sub to call an update to all existing graphs manually


```
Sub modeleren()

  runmodels ("all")

End Sub
```


Good Luck!


The MIT License (MIT)

Copyright (c) HVA 2018, Tony

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "modelerenclass" code), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

