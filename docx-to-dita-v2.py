from tkinter import Tk, Label, Button, Entry, filedialog
from tkinter import messagebox
from docx import Document
import xml.etree.ElementTree as ET
import xml.dom.minidom
import os

def docx_to_dita_task(docx_path, dita_path, task_id):
    doc = Document(docx_path)
    
    # Create the root element of the DITA task document with task_id attribute
    root = ET.Element('task', id=task_id)
    
    # Add the title to the DITA document
    title = ET.SubElement(root, 'title')
    title.text = doc.paragraphs[0].text  # Assuming the first paragraph is the title
    
    # Add task body content
    task_body = ET.SubElement(root, 'taskbody')
    
    # Process paragraphs to identify steps
    steps = ET.SubElement(task_body, 'steps')
    
    current_step = None
    
    for para in doc.paragraphs[1:]:  # Skipping the first paragraph (title)
        if para.style.name == 'List Paragraph':
            # New main step
            current_step = ET.SubElement(steps, 'step')
            step_cmd = ET.SubElement(current_step, 'cmd')
            step_cmd.text = para.text.strip()
        else:
            # Regular paragraphs are added as step info
            if current_step is not None:
                step_info = ET.SubElement(current_step, 'info')
                step_info.text = para.text.strip()
    
    # Create a new XML tree
    tree = ET.ElementTree(root)
    
    # Convert XML to string
    xml_str = ET.tostring(root, encoding='utf-8', method='xml')
    
    # Parse XML string for pretty printing
    dom = xml.dom.minidom.parseString(xml_str)
    pretty_xml_str = dom.toprettyxml(indent='  ')
    
    # Insert DOCTYPE declaration
    pretty_xml_str = '<?xml version="1.0" encoding="UTF-8"?>\n' \
                     '<!DOCTYPE task PUBLIC "-//OASIS//DTD DITA Task//EN" "task.dtd">\n' \
                     + pretty_xml_str
    
    # Write formatted XML to the .dita file
    with open(dita_path, 'w', encoding='utf-8') as f:
        f.write(pretty_xml_str)

def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    entry.delete(0, 'end')
    entry.insert(0, file_path)

def convert_file(input_path_entry, output_path_entry, task_id_entry):
    input_path = input_path_entry.get()
    output_path = output_path_entry.get()
    task_id = task_id_entry.get()
    
    if not input_path:
        messagebox.showerror("Error", "Please select an input .docx file.")
        return
    
    if not output_path:
        messagebox.showerror("Error", "Please specify an output .dita file.")
        return
    
    if not input_path.lower().endswith('.docx'):
        messagebox.showerror("Error", "Input file must be a .docx file.")
        return
    
    if not output_path.lower().endswith('.dita'):
        output_path += '.dita'
    
    docx_to_dita_task(input_path, output_path, task_id)
    messagebox.showinfo("Success", "Conversion completed successfully.")

def create_gui():
    window = Tk()
    window.title("Convert .docx to .dita")
    
    # Label and entry for input file path
    Label(window, text="Input .docx file:").grid(row=0, column=0)
    input_path_entry = Entry(window, width=50)
    input_path_entry.grid(row=0, column=1)
    Button(window, text="Browse", command=lambda: select_file(input_path_entry)).grid(row=0, column=2)
    
    # Label and entry for output file path
    Label(window, text="Output .dita file:").grid(row=1, column=0)
    output_path_entry = Entry(window, width=50)
    output_path_entry.grid(row=1, column=1)
    
    # Label and entry for task ID
    Label(window, text="Task ID:").grid(row=2, column=0)
    task_id_entry = Entry(window, width=50)
    task_id_entry.grid(row=2, column=1)
    
    # Convert button
    convert_button = Button(window, text="Convert", command=lambda: convert_file(input_path_entry, output_path_entry, task_id_entry))
    convert_button.grid(row=3, column=1)
    
    window.mainloop()

if __name__ == '__main__':
    create_gui()
