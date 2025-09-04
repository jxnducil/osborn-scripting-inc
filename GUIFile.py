from FunctionCodeForCETool import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkinter.messagebox import showinfo
import traceback

root = tk.Tk()


def UpdateText(String=" ", is_error=False):
    log_box.configure(state='normal')
    log_box.insert(tk.END, String + "\n")
    log_box.see(tk.END)  # auto-scroll
    log_box.configure(state='disabled')
    root.update_idletasks()  # forces GUI to update immediately

def browse_file():
    file_path= filedialog.askopenfilename(
        title="Select Civil 3D Excel File", 
        filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_path_var.set(file_path)
        file_label.config(text=f"Selected: {file_path}")

def run_estimation_gui():
    try:
        UpdateText("📂 Starting estimation...")
        result = CADToEstimate(file_path_var.get(),log=UpdateText,location=Selected_location.get())  # Your processing function
        UpdateText("✅ Done!")
    except Exception as e:
        UpdateText(f"❌ Error: {str(e)}", is_error=True)
        UpdateText(traceback.format_exc(), is_error=True)

root.title("Civil 3D Cost Estimation Tool")
root.geometry("1200x650")  # sets window size
root.configure(bg="orange2")

file_path_var = tk.StringVar()

instruction_label = tk.Label(root, text="Welcome to the Quick Pipe and Structure Tool. This tool was made with the intent to quicken and standardize the process of pipe network quanitifying and estimation." \
" In the near future more things will be implemented. Please share anything that you think would be benifital to add :)" , font=("Arial", 12,"bold"),bg="orange2",wraplength=1100, justify="center")
instruction_label.pack(pady=(10, 0))  # top padding 10px, bottom padding 0px

instruction_label2 = tk.Label(root, text="To get started, open one of the provided Excel templates and copy your CAD structure and pipe data into the appropriate sheets: PipeInput and StrucutreInput. Then click the Browse button below to locate the" \
"Excel file you just populated. Once seleceted, click the Run Button to beign the estimation process. The application will write the results directly into that same Excel file. Currently, the Standards dropdown doesn't change any calculations. All computations" \
" are based on WSDOT standards for now. However, the framework is in place to support additional standards, such as SPU, in the future. This dropdown will eventually allow you to choose your preferred standard. The Message box blwo will display any errors or improtant " \
"information that comes up during the estimation process", font=("Arial", 10),bg="orange2",wraplength=1000, justify="center")
instruction_label2.pack(pady=(14, 0))  # top padding 10px, bottom padding 0px


browse_button = tk.Button(root, text="Browse Excel File", command=browse_file)
browse_button.pack(pady=5)

file_label = tk.Label(root, text="No file selected.", bg="orange2")
file_label.pack(pady=5)

run_button = tk.Button(root, text="Run Estimation", command=run_estimation_gui)
run_button.pack(pady=5)

#Location Combo Box
Location_Label = ttk.Label(text="Please select Standards", font=("Arial", 10) )
Location_Label.pack(padx=5)
Selected_location = tk.StringVar()
Location_cb = ttk.Combobox(root, textvariable=Selected_location, width = 23)
Location_cb['values'] = ['WSDOT', 'SPU']  # Set values on the combobox
Location_cb['state'] = 'readonly'        # Set state on the combobox
Location_cb.pack(padx=5)
def Location_change(event):
    """ Handle the month Changed event"""
    showinfo(
        title='Result',
        message=f'You selected {Selected_location.get()}'
    )
Location_cb.bind('<<ComboboxSelected>>',Location_change)

        
Location_button = tk.Button(root, text="Run Estimation", command=run_estimation_gui)
run_button.pack(pady=5)

log_box = tk.Text(root, height=30, width=100, wrap='word')
log_box.pack(pady=30)

root.mainloop()
