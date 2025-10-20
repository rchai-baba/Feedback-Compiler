# %%
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# %%
def process_data(filename):
    responses = pd.read_csv(filename, skiprows=1, names=['Timestamp', 'Username', 'First Name', 'Last Name', 'Participation', 'Enthusiasm', 'Communication', 'Overall', 'Good', 'Improve'])
    responses = responses.drop(columns=['Timestamp','Username'])
    p_score = responses.groupby(['First Name', 'Last Name'])['Participation'].mean()
    e_score = responses.groupby(['First Name', 'Last Name'])['Enthusiasm'].mean()
    c_score = responses.groupby(['First Name', 'Last Name'])['Communication'].mean()
    o_score = responses.groupby(['First Name', 'Last Name'])['Overall'].mean()
    g_feedback = responses.groupby(['First Name', 'Last Name'])['Good'].agg(lambda x: "\n\t".join(x))
    b_feedback = responses.groupby(['First Name', 'Last Name'])['Improve'].agg(lambda x: "\n\t".join(x))
    final_df = pd.concat([p_score, e_score, c_score, o_score, g_feedback, b_feedback], axis=1)
    final_df.reset_index(inplace=True)
    print(final_df.head())
    return final_df

def write_data(output_df):
    with open('processed_feedback.txt', 'w') as output:
        for row in output_df.itertuples():
            output.write(str(row[1]) + ' ' + str(row[2]) + '\n\n')
            output.write('Participation: ' + str(row[3]) + '\n')
            output.write('Enthusiasm: ' + str(row[4]) + '\n')
            output.write('Communication: ' + str(row[5]) + '\n')
            output.write('Overall: ' + str(row[6]) + '\n')
            output.write('\n')
            output.write('Things you did well: \n')
            output.write('\t' + str(row[7]))
            output.write('\n')
            output.write('Things you could improve: \n')
            output.write('\t' + str(row[8]) + '\n\n\n')

# %%
def get_scaling_factor():
    """Get the DPI scaling factor"""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
        hdc = windll.user32.GetDC(0)
        dpi = windll.gdi32.GetDeviceCaps(hdc, 88)
        windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except:
        return 1.0

def browse_file():
    filename = filedialog.askopenfilename(
        title="Select feedback file",
        filetypes=[("CSV files", "*.csv")]
    )
    if filename:
        file_path_var.set(filename)

def submit():
    file_path = file_path_var.get()
   
    if not file_path:
        messagebox.showerror("Error", "Please select a file")
        return
   
    try:
        output_df = process_data(file_path)
        write_data(output_df)
        messagebox.showinfo("Success", "Processing complete! Check processed_feedback.txt")
        root.destroy()
    except Exception as e:
        messagebox.showerror("Error", f"Processing failed: {str(e)}")

# Create window
root = tk.Tk()
root.title("Feedback Compiler")

scale = get_scaling_factor()
width = int(450*scale)
height = int(100*scale)

# Get screen dimensions
root.geometry(f"{width}x{height}")
root.eval('tk::PlaceWindow . center')
# Variables
file_path_var = tk.StringVar()

# File selector row - no scale multiplication on padding
tk.Label(root, text="Feedback File:", font=("Arial", int(10*scale))).grid(row=0, column=0, sticky="w", padx=20, pady=15)
tk.Entry(root, textvariable=file_path_var, width=35, font=("Arial", int(10*scale))).grid(row=0, column=1, padx=10, pady=15)
tk.Button(root, text="Browse", command=browse_file, font=("Arial", int(10*scale))).grid(row=0, column=2, padx=10, pady=15)

# Submit button
tk.Button(root, text="Process Feedback", command=submit, bg="#4CAF50", fg="white", font=("Arial", int(12*scale), "bold"), padx=30, pady=10).grid(row=1, column=1, pady=20)

root.mainloop()