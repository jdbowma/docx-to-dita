# docx-to-dita
A Python script/app to convert .docx to DITA XML format. It is currently in its extremely early alpha phase and is considered extremely unreliable and unstable.

Currently only supports steps (including substeps) in a task topic. It relies upon steps being in a numbered (List Paragraph or List Number/List Number 2) list for now. 

Available in both GUI and CLI versions (CLI version lacks in features at the moment). 

Current features:
- Convert a .docx file to a DITA task topic nearly instantly (under 0.05 seconds)
- Automatic image handling (in early testing stages)
- Automatic note handling (any paragraph that starts with Note: will be turned into a <info><note> tag. There is an option to disable this and also an option to ask if each detected note is actually a note)
- Automatic keyword and phrase replacement (automatically detect terms and phrases that should be replaced with DITA keys, configurable from a preferences menu or from a preferences.json file. Works for other things too, not just keywords)

Planned features:
- Concept and reference topic types
- Automatic short descriptions
- Automatic fig/image captions
- Improved detection of lists and sub lists (not relying on the List Number/List Number 2 styles)

Known issues:
- Images output correctly, but are out of order and at the bottom of the generated DITA document
- Runs into permissions issues on Windows depending on where code is being run from, where preferences.json is saved, and where the .docx file is located

Installation:
1. Install Python 3
2. Install the python-docx module using pip (pip install python-docx)
3. Download the most recent, stable feature build of the docx-to-dita.py script and run it from a directory with properly configured permissions (the script needs to be able to read and write files)
4. Configure your keyword replacements from the settings window
5. Begin converting docx to DITA.
