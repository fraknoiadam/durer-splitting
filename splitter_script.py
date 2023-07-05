import os
import shutil
import webbrowser
from PyPDF2 import PdfWriter, PdfReader
import pandas as pd
from tkinter import *
from tkinter import messagebox, ttk
import pyautogui

# Parse the Excel file
df = pd.read_excel('XVI. Dürer Verseny - csapatadatok.xlsx')


# Create a dictionary of team names (as strings) and IDs
teams = {str(name): id for name, id in zip(df['Csapatnév'], df['ID'])}

longteams = {
    str(name): {
        'id': id,
        'category': category,
        'location': location,
        'first_member_name': first_member_name,
        'first_member_school': first_member_school
    } 
    for name, id, category, location, first_member_name, first_member_school 
    in zip(df['Csapatnév'], df['ID'], df['Kategória'], df['Helyszín'], df['1. tag neve'], df['1. tag iskolája'])
}

# Split PDF into separate pages
inputpdf = PdfReader(open("C5.pdf", "rb"))
for i in range(len(inputpdf.pages)):
    output = PdfWriter()
    output.add_page(inputpdf.pages[i])
    with open("document-page%s.pdf" % i, "wb") as outputStream:
        output.write(outputStream)


class AutocompleteEntry(ttk.Entry):
    def __init__(self, autocomplete_list, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._autocomplete_list = sorted(autocomplete_list)
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = StringVar()
        self.var.trace('w', self.changed)
        self.bind("<Return>", self.selection)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)
        self.listbox_up = False

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
        return [w for w in self._autocomplete_list if w.lower().startswith(self.var.get().lower())]

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



class PDFClassifier(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        # Create a list of team names, converting all to strings
        autocomplete_list = [str(name) for name in teams.keys()]
        self.entrythingy = AutocompleteEntry(autocomplete_list, master=root)
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
        if team_name in teams:
            src_file = f"document-page{self.page_counter}.pdf"
            team_id = teams[team_name]
            dst_file = os.path.join(str(team_id), src_file)
            # Check if the page file exists before trying to move it
            if os.path.isfile(src_file):
                if not os.path.exists(str(team_id)):
                    os.makedirs(str(team_id))
                # If the destination file already exists, don't try to move
                if not os.path.exists(dst_file):
                    shutil.move(src_file, dst_file)
                else:
                    print(f"File {dst_file} already exists")
            else:
                print(f"Page {self.page_counter} already classified")
            # Regardless of whether the file was moved, increment the counter and open the next page
            self.page_counter += 1
            self.entrythingy.after(1000, self.focus_entry)
            self.open_pdf()
        else:
            print("Team name not found.")

    

    def focus_entry(self):
        x, y, _, _ = self.entrythingy.bbox("end")
        x += self.entrythingy.winfo_rootx()
        y += self.entrythingy.winfo_rooty()
        pyautogui.click(x, y)
        pyautogui.press('right')  # simulating pressing the right arrow





    
        
            


root = Tk()
pdf_classifier = PDFClassifier(master=root)
pdf_classifier.mainloop()

