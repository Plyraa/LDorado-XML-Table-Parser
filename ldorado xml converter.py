import itertools
import os
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from xml.etree.ElementTree import parse

import pandas as pd
from openpyxl.workbook import Workbook


def select_xml_files():
    # Set default directory as C:\%USER%\Downloads for conventional purposes
    if os.name == "nt":
        initial_dir = f"{os.getenv('USERPROFILE')}\\Downloads"
    else:  # PORT: For *Nix systems
        initial_dir = f"{os.getenv('HOME')}/Downloads"

    file_types = (
        ('XML files', '*.xml'),
        ('All files', '*.*'),
    )

    # open-file dialog
    root = tk.Tk()
    filenames = tk.filedialog.askopenfilenames(
        initialdir=initial_dir,
        title='Select XML files',
        filetypes=file_types,
    )
    root.destroy()
    return filenames


def parse_xmls(selected_xml_files):
    excel_name = datetime.now().strftime("XML_%d-%m-%Y %H-%M-%S") + '.xlsx'
    try:
        excel_dir = os.path.join(os.path.dirname(selected_xml_files[0]), excel_name)
    except Exception as e:
        text.insert("0.0", "Error: Can't reach directory of selected XML file.")

    workbook = Workbook()
    ws1 = workbook.active
    ws1.title = "Sheet1"

    try:
        workbook.save(filename=excel_dir)
    except Exception as e:
        text.insert("0.0", "Error: Check permissions. Can't create Excel file at" + excel_dir + "\n")

    dataframe_list = []
    skipped_list = []
    for i, XMLfile in enumerate(selected_xml_files):
        # I'm not sure if this filename convention works for every output of LDorado…
        # It worked for my testing pool. Feel free to change.
        only_xml_name = XMLfile.split("/")
        only_xml_name = only_xml_name[len(only_xml_name) - 1].split("_")[0]

        try:
            tree = parse(XMLfile)
        except Exception as e:
            text.insert("0.0", XMLfile + " is not a valid XML file. Skipped." + "\n" + str(i + 1) + "/" + str(
                len(selected_xml_files)) + "\n")
            skipped_list.append(XMLfile)
            continue

        root = tree.getroot()
        tables_xpath = "./Harness/Tables/ComplexTable/SubTable"
        # row_xpath = "./Harness/Tables/ComplexTable/SubTable/Row"
        # data_xpath = "./Harness/Tables/ComplexTable/SubTable/Row/Cell/@Text"

        sub_tables = root.findall(tables_xpath)

        for table in sub_tables:
            table_rows = table.iterfind("./Row")
            splitted_cells = list(itertools.chain((it.findall("./Cell") for it in table_rows)))
            result_arrays = [[cell.get("Text") for cell in cells] for cells in splitted_cells]

            # If the table has only 1 element, silently ignore
            if len(result_arrays) < 2:
                continue

            filename_array = [only_xml_name] * (len(result_arrays) - 1)

            df = pd.DataFrame(result_arrays[1:], columns=result_arrays[0])
            df.insert(0, "Filename", filename_array)
            dataframe_list.append(df)

        text.insert("0.0",
                    "Reading tables from XML: " + str(XMLfile) + "\n" + str(i + 1) + "/" + str(
                        len(selected_xml_files)) + "\n")
        gui_root.update()

    text.insert("0.0",
                "Writing to Excel file... It takes time, don't panic!\n\n\n")
    gui_root.update()

    with pd.ExcelWriter(excel_dir, mode='a', engine="openpyxl",
                        if_sheet_exists="overlay") as writer:
        for dataf in dataframe_list:
            dataf.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row + 1, index=False)

    if len(skipped_list) > 0:
        text.insert("0.0",
                    "Corrupted XML file(s):" + str(skipped_list) + "\n\n\n")
        gui_root.update()

    text.insert("0.0",
                "Completed. Saved at " + excel_dir + "\n\n\n")
    gui_root.update()


def on_button_click():
    filelist = select_xml_files()
    text.delete("1.0", "end")
    start_time = time.time()
    parse_xmls(filelist)
    end_time = time.time()
    elapsed_time = end_time - start_time
    messagebox.showinfo("Conversion completed",
                        f"Time elapsed: {elapsed_time:.1f} seconds\n{len(filelist)} XML(s) processed", )
    # gui_root.destroy()


gui_root = tk.Tk()

gui_root.title("LDorado XML Converter")
gui_root.geometry("800x400")

label = tk.Label(gui_root, text="Output Log")
label.pack()

text = tk.Text(gui_root, wrap="word")
text.pack(fill="both", expand=True)

scrollbar = tk.Scrollbar(text, command=text.yview)
scrollbar.pack(side="right", fill="y")

text.config(yscrollcommand=scrollbar.set)
bottom_frame = tk.Frame(gui_root)
bottom_frame.pack(side="bottom", fill="x")

button = tk.Button(bottom_frame, text="Convert XMLs", command=on_button_click, width=20, height=6)
button.pack(pady=10)

gui_root.mainloop()
