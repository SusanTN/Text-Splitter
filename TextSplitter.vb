Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Drawing

Public Class TextSplitter
	Inherits Form
	Private WithEvents TextInput As TextBox
	Private WithEvents FooterText As TextBox
	Private WithEvents SplitText As Button
	Private WithEvents MaxCharacters As NumericUpDown
	Private WithEvents IncludeFooterInAllSegmentsCheckbox As CheckBox
	Private FooterTextLabel As Label
	Private maxCharactersLabel As Label
	Private totalCharactersLabel As Label
	Private numberOfSegmentsLabel As Label
	Private TextOutput As TextBox

	Public Sub New()
		InitializeComponent()
		CustomInitializeComponent()
	End Sub
	Protected Overrides Sub OnLoad(e As EventArgs)
		MyBase.OnLoad(e)
		AddHandler Me.Resize, AddressOf TextSplitter_Resize
	End Sub
	Private Sub CustomInitializeComponent()
		TextInput = New TextBox With {
			.Multiline = True,
			.Location = New Point(10, 10),
			.Size = New Size(300, 575),
			.MinimumSize = New Size(300, 575),
			.ScrollBars = ScrollBars.Vertical
		}
		TextOutput = New TextBox With {
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
		Size = New Size(642, 820)
		Text = "Text Splitter"
	End Sub
	Private Sub TextSplitter_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
		ApplyLayout() ' This will adjust the layout based on the current form size
	End Sub
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
	Private Sub SplitText_Click(sender As Object, e As EventArgs) Handles SplitText.Click
		Dim inputText As String = TextInput.Text
		Dim maxCharLimit As Integer = Convert.ToInt32(MaxCharacters.Value)
		Dim footerTemplate As String = FooterText.Text & " ({0}/{1})"
		Dim includeFooterInAllSegments As Boolean = IncludeFooterInAllSegmentsCheckbox.Checked

		Dim segments As List(Of String) = SplitIntoSegments(inputText, maxCharLimit, footerTemplate, includeFooterInAllSegments)
		TextOutput.Text = String.Join(Environment.NewLine, segments)

		totalCharactersLabel.Text = $"Total Characters: {inputText.Length}"
		numberOfSegmentsLabel.Text = $"Number of Segments: {segments.Count}"
	End Sub
	Private Function SplitIntoSegments(inputText As String, maxCharacters As Integer, footerText As String, includeFooterInAllSegments As Boolean) As List(Of String)
		' Dictionary to hold URL tags and their corresponding URLs
		Dim urlDictionary As New Dictionary(Of String, String)()
		Dim urlTagCounter As Integer = 0

		' Regex to find URLs
		Dim urlPattern As String = "https?://\S+"
		Dim regex As New Regex(urlPattern)

		' Replace URLs with tags and store them in the dictionary
		For Each match As Match In regex.Matches(inputText)
			Dim urlTag As String = $"[URL{urlTagCounter}]"
			urlDictionary.Add(urlTag, match.Value)
			inputText = inputText.Replace(match.Value, urlTag)
			urlTagCounter += 1
		Next

		Dim segments As New List(Of String)()
		Dim segmentCounter As Integer = 1

		While inputText.Length > 0
			Dim segmentSize As Integer = maxCharacters
			If segmentCounter = 1 OrElse includeFooterInAllSegments Then
				segmentSize -= footerText.Length + 10 ' Adjust size for footer
			End If

			Dim segmentEnd As Integer = Math.Min(segmentSize, inputText.Length)
			Dim segmentText As String = inputText.Substring(0, segmentEnd)

			' Find the last sentence terminator to avoid splitting mid-sentence
			Dim sentenceTerminators As Char() = {"."c, "!"c, "?"c, """"c, ChrW(&H201D)}
			Dim lastSentenceEnd As Integer = segmentText.LastIndexOfAny(sentenceTerminators)
			If lastSentenceEnd > -1 AndAlso lastSentenceEnd < segmentSize - 1 Then
				segmentText = segmentText.Substring(0, lastSentenceEnd + 1)
			End If

			' Append the footer or segment count as required, with formatting
			If segmentCounter = 1 OrElse includeFooterInAllSegments Then
				Dim footer As String = String.Format(footerText, segmentCounter, "{totalSegments}")
				segments.Add(segmentText.Trim() & Environment.NewLine & Environment.NewLine & "— " & footer & Environment.NewLine & Environment.NewLine)
			Else
				segments.Add(segmentText.Trim() & Environment.NewLine & Environment.NewLine & "— (" & segmentCounter.ToString() & "/{totalSegments})" & Environment.NewLine & Environment.NewLine)
			End If

			inputText = inputText.Substring(segmentText.Length).Trim()
			segmentCounter += 1
		End While

		' Replace the URL tags with actual URLs in each segment
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