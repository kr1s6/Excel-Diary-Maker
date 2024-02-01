
from tkinter import *
from tkinter import filedialog
from excel_maker import *

Y = 120


def browse_excel_file():
    filename = filedialog.askopenfilename(initialdir="/Desktop",
                                          title="Select an Excel file",
                                          filetypes=(("Excel files", "*.xlsx*"),
                                                     ("All files", "*.*")))
    excel_name.delete(0, END)
    excel_name.insert(0, filename)


def browse_notes_file():
    filename = filedialog.askopenfilename(initialdir="/Desktop",
                                          title="Select an txt file",
                                          filetypes=(("Text files", "*.txt*"),
                                                     ("All files", "*.*")))
    notes_name.delete(0, END)
    notes_name.insert(0, filename)


def update_create_diary(excel, notes):
    excel_maker(excel, notes)


if __name__ == '__main__':
    app = Tk()
    app.title("Excel Diary Maker")
    app.geometry('600x400')
    app.configure(background="white")
    text = Text(app, height=6, background='ghostwhite')
    text.insert(INSERT, "Program only works when notes are separated by 'Day (data) (title)'\n"
                        "Day 27.10.2023 Friend's birthday party\n"
                        "(body)\n"
                        "Day 28.10.2023 Campfire with friends\n"
                        "(body)")
    text.grid(column=0, row=0)

    # ------------------------Excel file----------------------------------------
    excel_label = Label(app, text="Path/name to the Excel you want to update/create:",
                        justify="left", background="white")
    excel_label.place(x=10, y=Y)

    excel_name = Entry(app, width=50, background="ghostwhite")
    excel_name.insert(0, "Journey.xlsx")
    excel_name.place(x=15, y=Y + 30)
    button_excel = (Button(app, text="Browse File", command=browse_excel_file)
                    .place(x=320, y=Y + 26))

    # ------------------------------Notes file------------------------------------
    notes_label = Label(app, text="Browse path to your notes (like '../Journey.txt'):",
                        justify="left", background="white")
    notes_label.place(x=10, y=Y + 80)

    notes_name = Entry(app, width=50, background="ghostwhite")
    notes_name.insert(0, "October_diary.txt")
    notes_name.place(x=15, y=Y + 110)
    button_notes = (Button(app, text="Browse File", command=browse_notes_file)
                    .place(x=320, y=Y + 106))

    # -------------------------------Final button-----------------------------------
    button_execute = Button(app, text="Update/Create your Diary",
                            command=lambda: update_create_diary(excel_name.get(), notes_name.get()))
    button_execute.place(x=430, y=Y + 230)
    # ------------------------------------------------------------------------------
    notes_label = Label(app, text="App 02.02.2024 Krzysztof B.",
                        justify="left", background="white")
    notes_label.place(x=10, y=Y + 260)

    app.mainloop()
