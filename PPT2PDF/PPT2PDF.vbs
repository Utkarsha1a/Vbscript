Create_PDF()
Dim inputpath, outputpath

Sub ConvertToPDF(ppt, inputpath, outputpath)
   Dim presentation
   Dim printoptions

   ppt.Presentations.Open inputpath

   set presentation = ppt.ActivePresentation
   set printoptions = presentation.PrintOptions

   printoptions.Ranges.Add 1,presentation.Slides.Count
   printoptions.RangeType = 1 ' Show all.

   const ppFixedFormatTypePDF = 2
   const ppFixedFormatIntentScreen = 1
   const msoFalse = 0
   const msoTrue = -1
   const ppPrintHandoutHorizontalFirst = 2
   const ppPrintOutputSlides = 1
   const ppPrintAll = 1
	 msgBox(inputFile)
   presentation.ExportAsFixedFormat outputpath, ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoTrue, ppPrintHandoutHorizontalFirst, ppPrintOutputSlides, msoFalse, printoptions.Ranges(1), ppPrintAll, inputFile, False, False, False, False, False

   presentation.Close
   msgBox(inputFile)
End Sub


 
Sub Create_PDF()

	set FSO = CreateObject("Scripting.FileSystemObject")
	set ppt = CreateObject("PowerPoint.Application")
	ppt.Visible = True

	  inputFile = "Input.pptx"
	If inputFile <> "" Then
	  If Not FSO.FileExists( inputFile ) Then
		 WScript.Stdout.Writeline "File not found: " & inputFile
	  End If

	  inputpath = FSO.GetAbsolutePathName(inputFile)
	  msgBox(inputpath)

	  outputpath = "Output.pdf"
		msgBox(outputpath)

	  ConvertToPdf ppt, inputpath, outputpath
	
   ppt.Quit
End If
msgbox("PDF Created")
End Sub