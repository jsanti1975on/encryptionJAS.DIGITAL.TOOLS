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
