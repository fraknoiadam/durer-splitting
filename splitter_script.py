import os
import shutil
import webbrowser
from PyPDF2 import PdfWriter, PdfReader
import pandas as pd
from tkinter import *
from tkinter import messagebox, ttk
import pyautogui

# Parse the Excel file
teams_df = pd.read_excel('XVI. Dürer Verseny - csapatadatok.xlsx')

# Split PDF into separate pages
inputpdf = PdfReader(open("C5.pdf", "rb"))
for i in range(len(inputpdf.pages)):
    output = PdfWriter()
    output.add_page(inputpdf.pages[i])
    with open("document-page%s.pdf" % i, "wb") as outputStream:
        output.write(outputStream)


class AutocompleteEntry(ttk.Entry):
    def __init__(self, master, autocomplete_list, **kwargs):
        super().__init__(master, **kwargs)
        self._autocomplete_list = sorted(autocomplete_list.keys())
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()
        self.var.trace('w', self.changed)
        self.bind("<Return>", self.selection)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.listbox_up = False
        self.master = master

    def changed(self, name, index, mode):
        if self.var.get() == '':
            if self.listbox_up:
                self.listbox.destroy()
                self.listbox_up = False
        else:
            words = self.comparison()
            if words:
                if not self.listbox_up:
                    self.listbox = Listbox()
                    self.listbox.bind("<Double-Button-1>", self.selection)
                    self.listbox.bind("<Right>", self.selection)
                    self.listbox.place(x=self.winfo_x(), y=self.winfo_y()+self.winfo_height())
                    self.listbox_up = True
                self.listbox.delete(0, END)
                for word in words:
                    self.listbox.insert(END, word)
            else:
                if self.listbox_up:
                    self.listbox.destroy()
                    self.listbox_up = False

    def comparison(self):
        # When isfull is True, include additional details in the autocomplete suggestions
        if self.master.isfull:
            return [f"{name} ({team_info['category']}, {team_info['location']}, {team_info['first_member_name']}, {team_info['first_member_school']})"
                    for name, team_info in self.master._autocomplete_list.items()
                    if name.lower().startswith(self.var.get().lower())]
        else:
            # When isfull is False, only include the team name in the autocomplete suggestions
            return [name for name in self._autocomplete_list if name.lower().startswith(self.var.get().lower())]


    def selection(self, event):
        if self.listbox_up:
            self.var.set(self.listbox.get(ACTIVE))
            self.listbox.destroy()
            self.listbox_up = False
            self.icursor(END)

    def up(self, event):
        if self.listbox_up:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]
            if index != '0':
                self.listbox.selection_clear(first=index)
                index = str(int(index)-1)
                self.listbox.see(index)
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)

    def down(self, event):
        if self.listbox_up:
            if self.listbox.curselection() == ():
                index = '0'
            else:
                index = self.listbox.curselection()[0]
            if index != END:
                self.listbox.selection_clear(first=index)
                index = str(int(index)+1)
                self.listbox.see(index)
                self.listbox.selection_set(first=index)
                self.listbox.activate(index)



def construct_autocomplete_list(teams_df):
    autocomplete_list = {}
    for _, row in teams_df.iterrows():
        team_name = row['Csapatnév']
        category = row['Kategória']
        location = row['Helyszín']
        first_member_name = row['1. tag neve']
        first_member_school = row['1. tag iskolája']
        id = row['ID']
        
        autocomplete_list[str(team_name)] = {
            'category': category,
            'location': location,
            'first_member_name': first_member_name,
            'first_member_school': first_member_school,
            'id': id
        }
    return autocomplete_list



class PDFClassifier(Frame):
    def __init__(self, master, teams_df):
        super().__init__(master)
        self.master = master
        self._autocomplete_list = construct_autocomplete_list(teams_df)
        self.entrythingy = AutocompleteEntry(self, self._autocomplete_list)
        self.entrythingy.focus_set()
        self.entrythingy.pack()
        self.contents = self.entrythingy.var
        self.entrythingy.bind('<Key-Return>', self.print_contents)
        self.page_counter = 0
        self.isfull = False  # Define variable to be changed
        self.button = ttk.Button(self, text="Csak név", command=self.on_button_click)
        self.button.pack()
        self.open_pdf()

    def on_button_click(self):
        self.isfull = not self.isfull  # Toggle the boolean variable on button click
        # Update the button's text based on the value of self.isfull
        self.button['text'] = "Teljes infó" if self.isfull else "Csak név"

    def open_pdf(self):
        webbrowser.open(f"document-page{self.page_counter}.pdf")
        
    def print_contents(self, event):
        team_name = self.contents.get()
        if team_name in self._autocomplete_list:
            src_file = f"document-page{self.page_counter}.pdf"
            team_id = self._autocomplete_list[team_name]['id']

            # Check if the page file exists before trying to move it
            if os.path.isfile(src_file):
                # Create the team folder if it doesn't exist
                team_folder = f"{team_id}"
                if not os.path.exists(team_folder):
                    os.makedirs(team_folder)
                
                # Try moving the file to the team folder. If a file with the same name already exists in the destination,
                # ignore the error and continue to the next file.
                dst_file = f"{team_folder}/document-page{self.page_counter}.pdf"
                if not os.path.isfile(dst_file):
                    shutil.move(src_file, dst_file)
                
                # Open the next page in the PDF
                self.page_counter += 1
                self.entrythingy.after(1000, self.focus_entry)
                self.open_pdf()
            else:
                print(f"Page {self.page_counter} already classified")
        else:
            print("Team name not found.")


    

    def focus_entry(self):
        x, y, _, _ = self.entrythingy.bbox("end")
        x += self.entrythingy.winfo_rootx()
        y += self.entrythingy.winfo_rooty()
        pyautogui.click(x, y)
        pyautogui.press('right')  # simulating pressing the right arrow





    
        
            


root = Tk()
pdf_classifier = PDFClassifier(root, teams_df)
pdf_classifier.pack()
pdf_classifier.mainloop()

