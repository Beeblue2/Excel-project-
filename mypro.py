from tkinter import *
from tkinter import ttk, filedialog
import numpy
import pandas as pd

root = Tk()
root.configure(bg="grey")
root.title("EXCEL SHEETS")
root.geometry("500x800")


# function for opening the file1
def open_file1():
    filename = filedialog.askopenfilename(title="Open the first file", filetype=(("xlxs files", ".*xlsx"),
                                                                                 ("All Files", "*.")))

    if filename:
        try:
            filename = r"{}".format(filename)
            global df1
            df1 = pd.read_excel(filename)
        except ValueError:
            label.config(text="File could not be opened")
        except FileNotFoundError:
            label.config(text="File Not Found")
            # Clear all the previous data in tree
    clear_treeview()


# function for opening the file2


def open_file2():
    filename = filedialog.askopenfilename(title="Open the second file", filetype=(("xlxs files", ".*xlsx"),
                                                                                  ("All Files", "*.")))

    if filename:
        try:
            filename = r"{}".format(filename)
            global df2
            df2 = pd.read_excel(filename)
        except ValueError:
            label.config(text="File could not be opened")
        except FileNotFoundError:
            label.config(text="File Not Found")
            # Clear all the previous data in tree
    clear_treeview()


def merge_files():
    dataframe = [df1, df2]

    # Remove duplicates from the merged pandas dataframes.
    merge = pd.concat(dataframe).drop_duplicates()

    merge.to_excel("Merged.xlsx")


# Clear the Treeview Widget


def clear_treeview():
    tree.delete(*tree.get_children())


# Create a Treeview widget
tree = ttk.Treeview(root)

label = Label(root, text="MERGING TWO EXCEL SHEETS", width=28, height=1, font=("times", 65, "bold"),
              bg="dark grey", fg="black")
label.pack()

b1 = Button(root, text="Get Sheet A", width=10, height=1, font=("normal", 15, "bold"),
            bg="aqua", bd=10, foreground="blue", command=open_file1)
b1.pack(pady=20)

b2 = Button(root, text="Get Sheet B", width=10, height=1, font=("time", 15, "bold"),
            bg="aqua", bd=10, foreground="blue", command=open_file2)
b2.pack(pady=20)

b3 = Button(root, text="Get Merged Sheet", width=15, height=1, font=("normal", 15, "bold"),
            bg="aqua", bd=10, foreground="black", command=merge_files)
b3.pack(pady=20)

root.mainloop()
