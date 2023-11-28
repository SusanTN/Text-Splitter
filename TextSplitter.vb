Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel
Imports System.Reflection.Emit

' TextSplitter: A form-based application for splitting text into segments

Public Class TextSplitter
	Inherits Form

	' Declarations for UI elements

	Private WithEvents TextInput As RichTextBox
	Private TextOutput As RichTextBox
	Private WithEvents FooterText As TextBox
	Private WithEvents SplitText As Button
	Private WithEvents MaxCharacters As NumericUpDown
	Private WithEvents IncludeFooterInAllSegmentsCheckbox As CheckBox
	Private FooterTextLabel As Label
	Private maxCharactersLabel As Label
	Private totalCharactersLabel As Label
	Private numberOfSegmentsLabel As Label
	' Example limit of 2,000,000 characters

	Private Const MaxTextSize As Integer = 2000000


	' Constructor to initialize the form components

	Public Sub New()
		InitializeComponent()
		CustomInitializeComponent()
	End Sub
	' Event handler for form load
	Protected Overrides Sub OnLoad(e As EventArgs)
		MyBase.OnLoad(e)
		AddHandler Me.Resize, AddressOf TextSplitter_Resize
	End Sub
	' Additional component initialization and layout settings
	Private Sub CustomInitializeComponent()
		' Initializing and configuring the form fields
		TextInput = New RichTextBox With {
			.Multiline = True,
			.Location = New Point(10, 10),
			.Size = New Size(300, 575),
			.MinimumSize = New Size(300, 575),
			.ScrollBars = ScrollBars.Vertical
		}
		TextOutput = New RichTextBox With {
			.Multiline = True,
			.Location = New Point(320, 10),
			.Size = New Size(300, 760),
			.MinimumSize = New Size(300, 760),
			.ScrollBars = ScrollBars.Vertical
		}
		FooterTextLabel = New Label With {
			.Text = "Footer Text",
			.Location = New Point(10, 120),
			.Size = New Size(100, 20)
		}
		FooterText = New TextBox With {
			.Location = New Point(120, 120),
			.Size = New Size(190, 20)
		}
		maxCharactersLabel = New Label With {
			.Text = "Max Characters",
			.Location = New Point(10, 150),
			.Size = New Size(100, 20)
		}
		MaxCharacters = New NumericUpDown With {
			.Location = New Point(120, 150),
			.Size = New Size(100, 20),
			.Minimum = 1,
			.Maximum = 1000,
			.Value = 500
		}
		SplitText = New Button With {
			.Text = "Split the text",
			.Location = New Point(10, 180),
			.Size = New Size(300, 30)
		}
		totalCharactersLabel = New Label With {
			.Text = "Total Characters: 0",
			.Location = New Point(10, 220),
			.Size = New Size(300, 20)
		}
		numberOfSegmentsLabel = New Label With {
			.Text = "Number of Segments: 0",
			.Location = New Point(10, 250),
			.Size = New Size(300, 20)
		}
		IncludeFooterInAllSegmentsCheckbox = New CheckBox With {
			.Text = "Include Footer in All Segments",
			.Location = New Point(10, 280),
			.Size = New Size(300, 20),
			.Checked = False
		}

		' Adding Form Controls

		Controls.Add(TextInput)
		Controls.Add(FooterTextLabel)
		Controls.Add(FooterText)
		Controls.Add(IncludeFooterInAllSegmentsCheckbox)
		Controls.Add(maxCharactersLabel)
		Controls.Add(MaxCharacters)
		Controls.Add(SplitText)
		Controls.Add(totalCharactersLabel)
		Controls.Add(numberOfSegmentsLabel)
		Controls.Add(TextOutput)
		' Generating New window 
		Size = New Size(642, 820)
		Text = "Text Splitter"
	End Sub

	' Event to automatically adjust the layout when the form is resized

	Private Sub TextSplitter_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
		ApplyLayout() ' This will adjust the layout based on the current form size
	End Sub

	'Sub to adjust the layout to the Application Size

	Private Sub ApplyLayout()
		If TextInput IsNot Nothing AndAlso TextOutput IsNot Nothing Then
			' Calculate available width and height, keeping a margin
			Dim availableWidth As Integer = Me.ClientSize.Width - 20
			Dim availableHeight As Integer = Me.ClientSize.Height - 60 ' Adjust as needed for margins and other controls

			' Calculate the widths for the text boxes
			Dim textBoxWidth As Integer = Math.Max(TextInput.MinimumSize.Width, CInt(availableWidth * 0.45))
			Dim textBoxHeight As Integer = CInt((availableHeight - 120) * 0.5) ' Adjust as needed

			' Set the width and height of the TextInput
			TextInput.Width = textBoxWidth
			TextInput.Height = textBoxHeight

			' Position FooterTextLabel directly above FooterText
			FooterTextLabel.Location = New Point(10, TextInput.Bottom + 10)
			FooterTextLabel.Width = textBoxWidth

			' Position FooterText directly below FooterTextLabel
			FooterText.Location = New Point(10, FooterTextLabel.Bottom + 5)
			FooterText.Width = textBoxWidth

			' Position IncludeFooterInAllSegmentsCheckbox below FooterText
			IncludeFooterInAllSegmentsCheckbox.Location = New Point(10, FooterText.Bottom + 10)
			IncludeFooterInAllSegmentsCheckbox.Width = textBoxWidth

			' Position maxCharactersLabel and MaxCharacters below FooterText
			maxCharactersLabel.Location = New Point(10, IncludeFooterInAllSegmentsCheckbox.Bottom + 10)
			MaxCharacters.Location = New Point(120, IncludeFooterInAllSegmentsCheckbox.Bottom + 10)
			MaxCharacters.Width = textBoxWidth - 110

			' Position SplitText button below MaxCharacters
			SplitText.Location = New Point(10, MaxCharacters.Bottom + 10)
			SplitText.Width = textBoxWidth

			' Position totalCharactersLabel to the left below SplitText
			totalCharactersLabel.Location = New Point(10, SplitText.Bottom + 10)
			totalCharactersLabel.Width = textBoxWidth / 2 - 5

			' Position numberOfSegmentsLabel to the right, next to totalCharactersLabel
			numberOfSegmentsLabel.Location = New Point(totalCharactersLabel.Right + 10, SplitText.Bottom + 10)
			numberOfSegmentsLabel.Width = textBoxWidth / 2 - 5

			' Set the location, width, and height of the TextOutput
			TextOutput.Location = New Point(TextInput.Right + 10, 10)
			TextOutput.Width = availableWidth - textBoxWidth - 10
			TextOutput.Height = textBoxHeight * 2 + 20 ' Adjust as needed
		End If
	End Sub

	' Event handler for the 'Split the text' button click

	Private Sub SplitText_Click(sender As Object, e As EventArgs) Handles SplitText.Click
		Dim inputText As String = TextInput.Text
		Dim maxCharLimit As Integer = Convert.ToInt32(MaxCharacters.Value)
		Dim footerTemplate As String = FooterText.Text & " ({0}/{1})"
		Dim includeFooterInAllSegments As Boolean = IncludeFooterInAllSegmentsCheckbox.Checked

		' Check to see if TextInput exceeds a maximum size.

		If inputText.Length > MaxTextSize Then
			MessageBox.Show($"Text exceeds the maximum limit of {MaxTextSize} characters.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
			Return
		End If

		' Call the function to split TextInput into idividual segments based on the Max Characters.

		Dim segments As List(Of String) = SplitIntoSegments(inputText, maxCharLimit, footerTemplate, includeFooterInAllSegments)

		' Display the split segments in the TextOutput TextBox

		TextOutput.Text = String.Join(Environment.NewLine, segments)

		' Update the UI with formatted total characters and number of segments

		totalCharactersLabel.Text = $"Total Characters: {inputText.Length.ToString("N0")}"
		numberOfSegmentsLabel.Text = $"Number of Segments: {segments.Count.ToString("N0")}"

	End Sub

	' Function to split the input text into segments based on the specified criteria

	Private Function SplitIntoSegments(inputText As String, maxCharacters As Integer, footerText As String, includeFooterInAllSegments As Boolean) As List(Of String)

		' Dictionary to temporarily replace and store URLs in the text
		' This is done to prevent URLs from being split across segments

		Dim urlDictionary As New Dictionary(Of String, String)()
		Dim urlTagCounter As Integer = 0

		' Regular expression to identify URLs in the text

		Dim urlPattern As String = "https?://\S+"
		Dim regex As New Regex(urlPattern)

		' Replacing URLs with unique tags and storing the mappings in the dictionary
		' This process helps in reinserting the URLs into their original positions later

		For Each match As Match In regex.Matches(inputText)
			Dim urlTag As String = $"[URL{urlTagCounter}]"
			urlDictionary.Add(urlTag, match.Value)
			inputText = inputText.Replace(match.Value, urlTag)
			urlTagCounter += 1
		Next

		' List to hold the split segments of text

		Dim segments As New List(Of String)()
		Dim segmentCounter As Integer = 1

		' Loop to split the text into segments based on the maximum character limit

		While inputText.Length > 0

			' Calculating segment size; adjusting for footer text if needed

			Dim segmentSize As Integer = maxCharacters
			If segmentCounter = 1 OrElse includeFooterInAllSegments Then
				segmentSize -= footerText.Length + 10 ' Adjust size for footer
			End If

			' Determining the endpoint of the current segment

			Dim segmentEnd As Integer = Math.Min(segmentSize, inputText.Length)
			Dim segmentText As String = inputText.Substring(0, segmentEnd)

			' Finding the end of the last complete sentence to avoid splitting mid-sentence

			Dim sentenceTerminators As Char() = {"."c, "!"c, "?"c, """"c, ChrW(&H201D)}
			Dim lastSentenceEnd As Integer = segmentText.LastIndexOfAny(sentenceTerminators)
			If lastSentenceEnd > -1 AndAlso lastSentenceEnd < segmentSize - 1 Then
				segmentText = segmentText.Substring(0, lastSentenceEnd + 1)
			End If

			' Adding the footer or segment count to the segment text
			' Formatting is applied to include segment counters

			If segmentCounter = 1 OrElse includeFooterInAllSegments Then
				Dim footer As String = String.Format(footerText, segmentCounter, "{totalSegments}")
				segments.Add(segmentText.Trim() & Environment.NewLine & Environment.NewLine & "— " & footer & Environment.NewLine & Environment.NewLine)
			Else
				segments.Add(segmentText.Trim() & Environment.NewLine & Environment.NewLine & "— (" & segmentCounter.ToString() & "/{totalSegments})" & Environment.NewLine & Environment.NewLine)
			End If

			' Preparing the remaining text for the next iteration

			inputText = inputText.Substring(segmentText.Length).Trim()
			segmentCounter += 1
		End While

		' Replacing the URL tags with their actual URLs in each segment

		Dim totalSegments As Integer = segments.Count
		For i As Integer = 0 To segments.Count - 1
			For Each kvp As KeyValuePair(Of String, String) In urlDictionary
				segments(i) = segments(i).Replace(kvp.Key, kvp.Value)
			Next
			segments(i) = segments(i).Replace("{totalSegments}", totalSegments.ToString())
		Next
		Return segments
	End Function
End Class
