![ROT13 PROJECT IMAGE 1](https://github.com/user-attachments/assets/68b84c37-8eb8-4968-bbce-de20b7702557)
# VBA ROT13 Decoder with Salting Feature

This project is a simple Excel-based VBA tool for decoding ROT13-encoded text, salting the decoded text with specific characters, and reverting the salted string back to its original form.

## Features
- **ROT13 Decoder**: Decodes text using the ROT13 cipher.
- **Salting**: Adds specific characters (`&`, `!`, and `$`) to the decoded text at predefined intervals.
- **Salt Removal**: Removes the salt characters from the salted text to revert it back to the decoded string.

## Requirements
- Microsoft Excel with VBA enabled.

## How to Use

### 1. Setting Up the UserForm

To use this tool, you'll need to insert the following controls into the Excel VBA UserForm:
- **Text Boxes**:
  - `txtInput`: Input the ROT13-encoded string here.
  - `txtOutput`: Displays the decoded string after ROT13 decoding.
  - `txtSalted`: Displays the salted version of the decoded string.
  
- **Command Buttons**:
  - `cmdRotate`: Decodes the ROT13 text entered in `txtInput` and displays the result in `txtOutput`.
  - `cmdAddSalt`: Adds salt characters (`&`, `!`, `$`) to the decoded string in `txtOutput` and displays the salted text in `txtSalted`.
  - `cmdRemoveSalt`: Removes the salt characters from the salted text in `txtSalted` and reverts it back to the original decoded string, displaying the result in `txtOutput`.

### 2. Workflow

#### Step 1: Decode ROT13 Text
1. Input the ROT13-encoded string in the `txtInput` text box.
2. Click the `cmdRotate` button to decode the string using ROT13.
   - The decoded string will appear in the `txtOutput` text box.

#### Step 2: Add Salt
1. After decoding, click the `cmdAddSalt` button to apply the following salting rules:
   - Add `&` after every 3rd character.
   - Add `!` after every 2nd character.
   - Add `$` after every 6th character.
2. The salted version of the decoded string will appear in the `txtSalted` text box.

#### Step 3: Remove Salt
1. To revert the salted string back to the original decoded string, click the `cmdRemoveSalt` button.
2. The salt characters (`&`, `!`, `$`) will be removed, and the result will appear in the `txtOutput` text box.

## Example

### Input (ROT13 Encoded Text):
Gur cnffjbeq vf 7k16JArUVv5LxVuJfsSVdbbtaHGlw9D4

### Decoded (After Clicking `cmdRotate`):
The password is 7x16WNeHIi5YkIhWsfFIqoognUTyj9Q4


### Salted Output (After Clicking `cmdAddSalt`):
T!h&epas!sw&or!d is$7x!1&6W!Ne&HI!i5$Yk!Ih&wfs!FI&qoo!gnU&Ty!j9$Q4!


### Reverted Output (After Clicking `cmdRemoveSalt`):
The password is 7x16WNeHIi5YkIhWsfFIqoognUTyj9Q4


## Code Snippets

### ROT13 Decoding Function
```vba

Private Sub cmdRotate_Click()
    ' Decode the text from txtInput and display it in txtOutput
    txtOutput.Text = ROT13(txtInput.Text)
End Sub

Private Sub cmdAddSalt_Click()
    ' Add salt to the decoded text and display it in txtSalted
    txtSalted.Text = AddSalt(txtOutput.Text)
End Sub

Private Sub cmdRemoveSalt_Click()
    ' Call the RevertSaltedText function to remove salt and update txtOutput
    RevertSaltedText
End Sub

Function ROT13(sInput As String) As String
    Dim sOutput As String
    Dim i As Integer
    Dim ch As String
    Dim charCode As Integer
    
    sOutput = ""
    
    For i = 1 To Len(sInput)
        ch = Mid(sInput, i, 1)
        charCode = Asc(ch)
        
        ' Check if the character is uppercase
        If charCode >= 65 And charCode <= 90 Then
            ' Rotate the character by 13 places in the alphabet
            charCode = ((charCode - 65 + 13) Mod 26) + 65
            
        ' Check if the character is lowercase
        ElseIf charCode >= 97 And charCode <= 122 Then
            ' Rotate the character by 13 places in the alphabet
            charCode = ((charCode - 97 + 13) Mod 26) + 97
        End If
        
        sOutput = sOutput & Chr(charCode)
    Next i
    
    ROT13 = sOutput
End Function

Function AddSalt(decodedText As String) As String
    Dim saltedText As String
    Dim i As Integer
    Dim insertPos As Integer
    Dim addChar As String
    
    saltedText = ""
    
    For i = 1 To Len(decodedText)
        ' Append the current character from the decoded text
        saltedText = saltedText & Mid(decodedText, i, 1)
        
        ' Apply salt according to the position
        If i Mod 6 = 0 Then
            addChar = "$"
        ElseIf i Mod 3 = 0 Then
            addChar = "&"
        ElseIf i Mod 2 = 0 Then
            addChar = "!"
        Else
            addChar = ""
        End If
        
        ' Append the salted character (if any)
        saltedText = saltedText & addChar
    Next i
    
    AddSalt = saltedText
End Function

Function RemoveSalt(saltedText As String) As String
    Dim cleanText As String
    Dim i As Integer
    Dim currentChar As String
    
    cleanText = ""
    
    For i = 1 To Len(saltedText)
        currentChar = Mid(saltedText, i, 1)
        
        ' Only add the character to the cleanText if it is not one of the salt characters
        If currentChar <> "&" And currentChar <> "!" And currentChar <> "$" Then
            cleanText = cleanText & currentChar
        End If
    Next i
    
    RemoveSalt = cleanText
End Function

Sub RevertSaltedText()
    ' Remove the salt from the text in txtSalted and display the result in txtOutput
    txtOutput.Text = RemoveSalt(txtSalted.Text)
End Sub


