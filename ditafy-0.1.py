import logging
import tkinter as tk
from tkinter import filedialog, messagebox, Menu, Toplevel, ttk
import json
import xml.etree.ElementTree as ET
from docx import Document
import xml.dom.minidom
import os
from PIL import Image, ImageTk
import io
import time
import cProfile

logger = logging.getLogger(__name__)
logging.basicConfig(format='%(asctime)s %(levelname)s %(message)s')
logger.setLevel(logging.CRITICAL)

# Global variable for storing keyword replacements
keyword_replacements = {}
# Save 'preferences' function
def save_preferences(window):
    global keyword_replacements
    try:
        prefs = preferences_text.get("1.0", tk.END).strip().split('\n')
        keyword_replacements = {}
        for pref in prefs:
            if ':' in pref:
                original, new = map(str.strip, pref.split(':', 1))
                keyword_replacements[original] = new
        # Save preferences to preferences.json
        if not os.path.exists('docx-to-dita'):
            os.makedirs('docx-to-dita')                  
        with open('docx-to-dita/keywords.json', 'w') as f:
            json.dump(keyword_replacements, f)
        messagebox.showinfo("Preferences Saved", "Keyword replacements saved successfully.")
        logger.info(f"Keyword replacements saved to docx-to-dita/preferences.json")
        window.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving preferences:\n{str(e)}")
        logger.error(str(e))
        
# 'Preferences' GUI
def open_preferences_dialog():
    global preferences_text
    preferences_window = tk.Toplevel(root)
    window = preferences_window
    preferences_window.title("Keyword Replacements")
    
    tk.Label(preferences_window, text="Specify replacements in the format 'ORIGINALPHRASE : NEWPHRASE'").grid(row=0, column= 0, padx=10, pady=5)
    
    preferences_text = tk.Text(preferences_window, width=80, height=20)
    preferences_text.grid(row=1, column=0, padx=10, pady=5)
    
    current_prefs = '\n'.join([f"{key} : {value}" for key, value in keyword_replacements.items()])
    preferences_text.insert(tk.END, current_prefs)

    tk.Button(preferences_window, text="Save Preferences", command=lambda: save_preferences(window)).grid(row=2, column=0, padx=10, pady=10)

# Image preview GUI. Not working super great right now.
def preview_image(img_data, next_image_callback):
    preview_window = Toplevel(root)
    preview_window.title("Image Preview")
    
    image = Image.open(io.BytesIO(img_data))
    img_width, img_height = image.size
    logger.info(f"Previewed an image of size {img_width}, {img_height}")
    img_dimensions_label = tk.Label(preview_window, text=f"Size: {img_width}px x {img_height}px")
    img_dimensions_label.pack()
    
    image.thumbnail((img_width, img_height))
    img = ImageTk.PhotoImage(image)
    
    label = tk.Label(preview_window, image=img)
    label.image = img
    label.pack()

    skip_button = tk.Button(preview_window, text="Skip Image", command=lambda: next_image_callback(preview_window, False))
    skip_button.pack(side=tk.LEFT)
    
    save_button = tk.Button(preview_window, text="Save Image", command=lambda: next_image_callback(preview_window, True))
    save_button.pack(side=tk.RIGHT)
    
    return preview_window

# Save image function. Also not working great.
def save_image_with_preview(img_data):
    image_path = [None]
    
    def next_image_callback(window, save):
        if save:
            image_path[0] = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg"), ("All files", "*.*")])
            if image_path[0]:
                image = Image.open(io.BytesIO(img_data))
                image.save(image_path[0])
                logger.info(f"Saved an image as {image_path}")
        else:
            logger.info("Image skipped.")
        window.destroy()
    
    preview_window = preview_image(img_data, next_image_callback)
    root.wait_window(preview_window)
    return image_path[0]

# Main conversion function
def docx_to_dita_task(docx_path, dita_path, topic_id, window):
    # Load the input .docx file
    start_time = time.perf_counter()
    doc = Document(docx_path)
    logger.info(f"Loaded {docx_path}")
    total_pause_time = 0
    total_title_pause_time = 0
    total_note_pause_time = 0 
    progressbar = ttk.Progressbar(window, orient=tk.HORIZONTAL, length=300)
    progressbar.grid(row=7, column=1, padx=10, pady=5)
    
    root = ET.Element('task', id=topic_id)
    # Set the first line to be the title
    title = ET.SubElement(root, 'title')
    if doc.paragraphs[0].style.name.startswith("Heading"):
        title_pause_time = time.perf_counter()
        title_response = messagebox.askyesno("Title detected", f"Is this the title? {doc.paragraphs[0].text}")
        if title_response:
            if doc.paragraphs[0].text.startswith("Step"):
                if doc.paragraphs[0].text[6] == ":":
                    print('colon')
                    title.text = doc.paragraphs[0].text[8:]
                else:
                    title.text = doc.paragraphs[0].text[5:]
            else:
                title.text = doc.paragraphs[0].text
            logger.info(f"Set title to {title.text}")
        else:
            logger.info(f"Possible title {doc.paragraphs[0].text} was denied by the user.")
        title_resume_time = time.perf_counter()
        total_title_pause_time = title_resume_time - title_pause_time
        total_pause_time = total_title_pause_time
        # print(total_pause_time)
    # For now, regardless of style just set the first paragraph to be the title
    else:
        title.text = doc.paragraphs[0].text
    # Insert shortdesc after title if applicable
    shortdesc_added = False
    next_para = doc.paragraphs[1].text.strip()
    if doc.paragraphs[1].style.name not in ['List Paragraph', 'List Number', 'List Number 2']:
        shortdesc_pause_time = time.perf_counter()
        response = messagebox.askyesno("Short Description Detected", f"Is this the short description?\n\n{next_para}")
        logger.info(f"Short description detected: {next_para}")
        if response:
            shortdesc = ET.SubElement(root, 'shortdesc')
            shortdesc.text = next_para
            shortdesc_added = True
            logger.info(f"Added short description: {shortdesc.text}")
        else:
            logger.info(f"User determined that '{next_para}' was not the shortdesc")
        shortdesc_resume_time = time.perf_counter()
        total_shortdesc_pause_time = shortdesc_resume_time - shortdesc_pause_time
        total_pause_time = total_pause_time + total_shortdesc_pause_time
    task_body = ET.SubElement(root, 'taskbody')
    steps = ET.SubElement(task_body, 'steps')
    
    current_step = None
    current_substeps = None
    step_number = 0
    progress_tracker_full = len(doc.paragraphs)
    progress_increment = 100/progress_tracker_full
    def update_progress():
        progressbar.step(progress_increment)
        window.update_idletasks()
    # for para in doc.paragraphs:
    #     para_text = para.text.strip()
    #     if para_text.startswith(" "):
    #         del(para_text[0])
    #         print("did it")
    for para in doc.paragraphs[1:]:
        para_text = para.text.strip()
        update_progress()
        if shortdesc_added and para_text == next_para:
            logger.info(f"Skipped first paragraph as it is the short description")
            shortdesc_added = False  # Reset after skipping the paragraph
            continue
        # Check for notes
        if check_for_notes.get() and para_text.startswith("Note:"):
            note_content = para_text[5:].strip()
            if prompt_for_notes.get():
                note_pause_time = time.perf_counter()
                response = messagebox.askyesno("Note Detected", f"Is this a note?\n\n{note_content}")
                logger.info(f"Detected a possible note: {note_content}")
                if response:
                    note_tag = ET.Element('note')
                    note_tag.text = note_content
                    info_tag = ET.SubElement(current_step if current_step else steps, 'info')
                    info_tag.append(note_tag)
                    logger.info(f"Added note: {note_content}")
                else:
                    logger.info(f"Possible note ({note_content}) was determined by the user to not be a note.")
                    if current_step is not None:
                        step_info = ET.SubElement(current_step, 'info')
                        step_info.text = note_content
                    else:
                        para_tag = ET.SubElement(steps, 'info')
                        para_tag.text = note_content
                note_resume_time = time.perf_counter()
                total_note_pause_time = note_resume_time - note_pause_time
                total_pause_time = total_pause_time + total_note_pause_time
                # print(total_pause_time)
            else:
                note_tag = ET.Element('note')
                note_tag.text = note_content
                info_tag = ET.SubElement(current_step if current_step else steps, 'info')
                info_tag.append(note_tag)
        elif para_text == "" or para_text == " ":
            progressbar['value'] += progress_increment
            continue
        elif para.style.name in {'List Paragraph', 'List Number'}:
            # Main step conversion
            step_number += 1
            current_step = ET.SubElement(steps, 'step')
            step_cmd = ET.SubElement(current_step, 'cmd')
            step_cmd.text = para_text
            logger.info(f"Added '{step_cmd.text}' as step {step_number}.")
        elif para.style.name == 'List Number 2':
            # Substeps
            if current_step is not None:
                if current_substeps is None:
                    current_substeps = ET.SubElement(current_step, 'substeps')
                substep = ET.SubElement(current_substeps, 'substep')
                substep_cmd = ET.SubElement(substep, 'cmd')
                substep_cmd.text = para_text
                logger.info(f"Added '{substep_cmd.text}' as a substep under step {step_number}.")
        else:
            if current_step is not None:
                step_info = ET.SubElement(current_step, 'info')
                step_info.text = para_text
            else:
                para_tag = ET.SubElement(steps, 'info')
                para_tag.text = para_text
        #root.update_idletasks() 

    # Images

    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            if include_images.get():
                img = rel.target_part.blob
                image_path = save_image_with_preview(img)
                if image_path:
                    if current_step is not None:
                        info_tag = ET.SubElement(current_step, 'info')
                        fig_tag = ET.SubElement(info_tag, 'fig')
                        title_tag = ET.SubElement(fig_tag, 'title')
                        title_tag.text = ""
                        img_tag = ET.SubElement(fig_tag, 'image', href=image_path)
    # XML building
    xml_str = ET.tostring(root, encoding='utf-8', method='xml').decode('utf-8')
    for original, new in keyword_replacements.items():
        xml_str = xml_str.replace(original, new)
        logger.info(f"Replaced {original} with {new}")
    root = ET.fromstring(xml_str)
    
    tree = ET.ElementTree(root)
    xml_str = ET.tostring(root, encoding='utf-8', method='xml')
    dom = xml.dom.minidom.parseString(xml_str)
    pretty_xml_str = dom.toprettyxml(indent='  ')
    pretty_xml_str = '\n'.join(pretty_xml_str.split('\n')[1:])
    xml_declaration = '<?xml version="1.0" encoding="UTF-8"?>\n'
    doctype = '<!DOCTYPE task PUBLIC "-//OASIS//DTD DITA Task//EN" "task.dtd">\n'
    final_xml_str = xml_declaration + doctype + pretty_xml_str
    
    with open(dita_path, 'w', encoding='utf-8') as f:
        f.write(final_xml_str)
    logger.info("Built final XML tree.")
    end_time=time.perf_counter()
    print(f"Performance time: {len(doc.paragraphs)} paragraphs in {(end_time - total_pause_time) - start_time:.6f} seconds")
    progressbar.destroy()

# Function to open system dialog for input .docx file
def browse_file():
    filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if filename:
        input_path_entry.delete(0, tk.END)
        input_path_entry.insert(0, filename)

def browse_output_file():
    output_file = filedialog.asksaveasfilename(filetypes=[("DITA files", ".*dita")])
    if output_file:
        output_path_entry.delete(0, tk.END)
        output_path_entry.insert(0, output_file)

# Conversion function called on "Convert" button press. Name is confusing. 
def profile():
        input_path = input_path_entry.get()
        output_path = output_path_entry.get()
        topic_id = topic_id_entry.get()

        if not input_path:
            messagebox.showerror("Error", "Please select an input file.")
            logger.error("User did not specify an input file.")
            return
        
        if not output_path:
            messagebox.showerror("Error", "Please specify an output file.")
            logger.error("User did not specify an output file.")
            return
        
        if not topic_id:
            messagebox.showerror("Error", "Please enter a task ID.")
            logger.error("User did not specify a task ID.")
            return
        
        if not input_path.lower().endswith('.docx'):
            messagebox.showerror("Error", "Input file must be a .docx file.")
            logger.error("User did not specify a .docx file as the input file.")
            return
        
        if not output_path.lower().endswith('.dita'):
            output_path += '.dita'
        
        try:
            # Actually call the docx_to_dita_task function
            docx_to_dita_task(input_path, output_path, topic_id, root)
            messagebox.showinfo("Success", f"Conversion completed successfully. Output saved to {output_path}")
            logger.info(f"Conversion complete. Output saved to {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during conversion:\n{str(e)}")
            logger.error(f"An error occurred during conversion: {str(e)}")
def convert_file():
    cProfile.run('profile()')

# GUI
root = tk.Tk()
root.resizable(False, False)
root.title("DOCX to DITA Converter")

menu_bar = Menu(root)
root.config(menu=menu_bar)

file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open Preferences", command=open_preferences_dialog)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

input_path_label = tk.Label(root, text="Input .docx file:")
input_path_label.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)

input_path_entry = tk.Entry(root, width=50)
input_path_entry.grid(row=0, column=1, padx=10, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=5)

output_path_label = tk.Label(root, text="Output .dita file:")
output_path_label.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)

output_path_entry = tk.Entry(root, width=50)
output_path_entry.grid(row=1, column=1, padx=10, pady=5)

output_path_button = tk.Button(root, text="Browse", command=browse_output_file)
output_path_button.grid(row=1, column=2, padx=10, pady=5)

topic_type = "Task"
def on_select(event):
    topic_type = dropdown.get()

topic_type_label = tk.Label(root, text="Topic type:")
topic_type_label.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)

options = ["Task", "Concept", "Reference"]
dropdown = ttk.Combobox(root, values=options, state="readonly")
dropdown.set(options[0])
dropdown.grid(row=2, column=1, columnspan=1, padx=10, pady=5)
dropdown.bind("<<ComboboxSelected>>", on_select)

topic_id_label = tk.Label(root, text="Topic ID:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
topic_id_entry = tk.Entry(root, width=50)
topic_id_entry.grid(row=3, column=1, padx=10, pady=5)

include_images = tk.BooleanVar()
include_images_checkbutton = tk.Checkbutton(root, text="Include Images", variable=include_images)
include_images_checkbutton.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)

check_for_notes = tk.BooleanVar()
check_for_notes_checkbutton = tk.Checkbutton(root, text="Check for notes", variable=check_for_notes)
check_for_notes_checkbutton.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)

prompt_for_notes = tk.BooleanVar()
prompt_for_notes_checkbutton = tk.Checkbutton(root, text="Prompt for note verification", variable=prompt_for_notes)
prompt_for_notes_checkbutton.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky=tk.W)

# change to "convert_file" to see cProfile output
convert_button = tk.Button(root, text="Convert", command=profile)
convert_button.grid(row=8, column=1, padx=10, pady=10)

preferences_button = tk.Button(root, text="Preferences", command=open_preferences_dialog)
preferences_button.grid(row=8, column=2, padx=10, pady=10)

preferences_file = f'docx-to-dita/keywords.json'
# Load keyword replacement preferences
if os.path.exists(preferences_file):
    with open(preferences_file, 'r') as f:
        keyword_replacements = json.load(f)

root.mainloop()
