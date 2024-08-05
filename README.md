# DITAfy - Convert Microsoft Word (.docx) to Software Documentation Using DITA
![DITAfy logo](https://github.com/jdbowma/docx-to-dita/blob/main/ditafy-removebg-preview.png)
#### DITAfy Docx-to-DITA is a GUI Python app that converts Microsoft Word '.docx' files into DITA XML for use in documentation authoring applications like Oxygen XML Author. It is currently in its very early alpha phase and should be considered unreliable.

Available in both GUI and CLI versions (CLI version lacks in features at the moment). 

### Current features:
- Convert a .docx file to a DITA task topic nearly instantly. Conversion times are typically faster than Oxygen's MS Word to DITA conversion. Currently only supports task topics. It relies upon steps being in a numbered (List Paragraph or List Number/List Number 2 list for now
- Automatic image handling (in early testing stages)
- Automatic note handling (any paragraph that starts with Note: will be turned into a <info><note> tag. There is an option to disable this and also an option to ask if each detected note is actually a note)
- Automatic keyword and phrase replacement (automatically detect terms and phrases that should be replaced with DITA keys, configurable from a preferences menu or from a preferences.json file. Works for other things too, not just keywords)
- Automatic short descriptions and titles (user is prompted to confirm title and short description)

### Planned features:
- Concept and reference topic types
- Automatic fig/image captions
- Better image handling
- Tables
- Optimized algorithms
- Improved detection of lists and sub-lists (not relying on the List Number/List Number 2 styles)

### Known issues:
- Images output correctly, but are out of order and at the bottom of the generated DITA document
- Runs into permissions issues on Windows depending on where code is being run from, where preferences.json is saved, and where the .docx file is located

### Installation:
1. Install Python 3
2. Install the python-docx module using pip (pip install python-docx)
3. Download the most recent, stable feature build of the docx-to-dita.py script and run it from a directory with properly configured permissions (the script needs to be able to read and write files)
4. Configure your keyword replacements from the settings window
5. Begin converting docx to DITA.

### Usage
1. Click "Browse" next to the "Input .docx file" field and select a properly styled and formatted Microsoft Word (.docx) file.
2. In the "Output .dita file" field, either specify the name of the output file manually or press "Browse" to select a directory and name the file from the system prompt.
3. Specify a topic ID.
4. Check the "Check for Notes", "Prompt for Notes", and/or "Include Images" checkboxes depending on your needs.
5. Click "Preferences" to configure keyword replacements.
6. Press "Convert". Check the console for any errors.
