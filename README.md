# searchBlue

A small Microsoft Word VBA macro that finds the next occurrence of
text formatted with a specific blue font colour.
This can be useful when documentation changes are marked up with a particular colour and you want to cycle through all changes.


## What searchBlue does

- Opens the Word Document Map
- Searches for text using a specific font colour
- Starts from the current cursor position
- Continues searching from the beginning when it reaches the end, looping around.

## Installation

### Option 1: Import the VBA module

1. Download `searchBlue.bas` from this repository.
2. Open Microsoft Word.
3. Press `Alt+F11` to open the Visual Basic Editor.
4. Select **File → Import File...**
5. Select `searchBlue.bas`.

The macro will then appear in the VBA project.

### Option 2: Copy the code

Open the VBA editor with `Alt+F11`, create a new module, and paste
the contents of `searchBlue.bas` into it.

## Usage

Run:
Make a macro button for the title bar and click.
