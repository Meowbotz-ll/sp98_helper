import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def parse_sort_order(input_str):
    try:
        classes = [s.strip() for s in input_str.strip().split('\n') if s.strip()]
        return {classname: index + 1 for index, classname in enumerate(classes)}
    except Exception as e:
        messagebox.showerror("Error", f"Invalid sort order format: {e}")
        return None

def toggle_sort_order_visibility():
    if sort_by_score_var.get():
        sort_order_frame.pack_forget()
    else:
        sort_order_frame.pack(fill='x')

def update_sort_order_list():
    if not file_path:
        return
    try:
        df = pd.read_excel(file_path)
        selected_column = sort_order_column_combobox.get()
        if selected_column in df.columns:
            values = df[selected_column].dropna().unique()
            if all(isinstance(value, (int, float)) for value in values):
                sort_by_score_var.set(True)
                toggle_sort_order_visibility()
            else:
                non_numerical_values = [value for value in values if not isinstance(value, (int, float))]
                sort_order_text.delete("1.0", tk.END)
                sort_order_text.insert("1.0", "\n".join(str(value) for value in non_numerical_values))
                sort_by_score_var.set(False)
                toggle_sort_order_visibility()
    except Exception as e:
        messagebox.showerror("Error", f"Unable to update sort order list: {e}")

def add_additional_column_selector():
    frame = ttk.Frame(additional_columns_frame, padding="3")
    frame.pack(fill='x')

    combobox = ttk.Combobox(frame)
    combobox['values'] = additional_columns_comboboxes[0]['values'] if additional_columns_comboboxes else []
    combobox.pack(side='left', fill='x', expand=True)

    additional_columns_comboboxes.append(combobox)

def process_file():
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return

    selected_column = sort_order_column_combobox.get()
    if not selected_column:
        messagebox.showerror("Error", "No column selected for sorting!")
        return

    selected_name_column = name_column.get()
    if not selected_name_column:
        messagebox.showerror("Error", "No name column selected!")
        return

    additional_info_columns = [combobox.get() for combobox in additional_columns_comboboxes if combobox.get()]

    try:
        # Parse group sizes
        group_sizes = [int(size.strip()) for size in group_size_entry.get().split(',')]

        if sort_by_score_var.get():
            df_sorted = pd.read_excel(file_path).sort_values(by=selected_column, ascending=False)
        else:
            custom_sort_order = parse_sort_order(sort_order_text.get("1.0", "end-1c"))
            if custom_sort_order is None:
                return
            df = pd.read_excel(file_path)
            df['sort_key'] = df[selected_column].map(custom_sort_order)
            df_sorted = df.sort_values(by='sort_key')

        start_index = 0
        for group_index, size in enumerate(group_sizes):
            group_label = 'Group' + chr(65 + group_index)
            file_name = f'{group_label}.txt'
            with open(file_name, 'w', encoding='GB18030') as file:
                for index in range(start_index, min(start_index + size, len(df_sorted))):
                    line = df_sorted.iloc[index][selected_name_column]
                    for col in additional_info_columns:
                        line += f'\t{df_sorted.iloc[index][col]}'
                    file.write(f'{line}\n')
            start_index += size

            # If the last group size has been reached, all remaining groups will have this size
            if group_index == len(group_sizes) - 1:
                while start_index < len(df_sorted):
                    group_index += 1
                    group_label = 'Group' + chr(65 + group_index)
                    file_name = f'{group_label}.txt'
                    with open(file_name, 'w', encoding='GB18030') as file:
                        for index in range(start_index, min(start_index + size, len(df_sorted))):
                            line = df_sorted.iloc[index][selected_name_column]
                            for col in additional_info_columns:
                                line += f'\t{df_sorted.iloc[index][col]}'
                            file.write(f'{line}\n')
                    start_index += size

        messagebox.showinfo("Success", "Files processed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        file_label.config(text="File Selected: " + file_path)
        update_columns()

def update_columns():
    try:
        df = pd.read_excel(file_path)
        columns = list(df.columns)
        sort_order_column_combobox['values'] = columns
        name_column_combobox['values'] = columns
        for combobox in additional_columns_comboboxes:
            combobox['values'] = columns
    except Exception as e:
        messagebox.showerror("Error", f"Unable to read file: {e}")

# GUI setup
root = tk.Tk()
root.title("sp98helper")
root.geometry("650x350")  # Adjust window size as needed

# Global variables
file_path = ''
sort_by_score_var = tk.BooleanVar(value=False)
group_size_entry = tk.StringVar(value="8")
name_column = tk.StringVar(value='')  # Initially empty
additional_columns_comboboxes = []

# Browse File Section (spanning both columns)
frame_file = ttk.Frame(root, padding="10")
frame_file.grid(row=0, column=0, columnspan=2, sticky='ew')

browse_button = ttk.Button(frame_file, text="Browse", command=browse_file)
browse_button.pack(side='left')

file_label = ttk.Label(frame_file, text="No file selected")
file_label.pack(side='left', padx=10)

# Left Column
left_column_frame = ttk.Frame(root, padding="10")
left_column_frame.grid(row=1, column=0, sticky='nsew', padx=5)

# Sort By Score Checkbox (left column)
frame_sort_by_score = ttk.Frame(left_column_frame, padding="10")
frame_sort_by_score.pack(fill='x')

sort_by_score_checkbox = ttk.Checkbutton(frame_sort_by_score, text="Sort By Score(Descending Order)", variable=sort_by_score_var, command=toggle_sort_order_visibility)
sort_by_score_checkbox.pack(side='left')

# Sort Order Column Selection Section (left column)
frame_sort_order_column = ttk.Frame(left_column_frame, padding="10")
frame_sort_order_column.pack(fill='x')

sort_order_column_label = ttk.Label(frame_sort_order_column, text="Sort Order Column:")
sort_order_column_label.pack(side='left')

sort_order_column_combobox = ttk.Combobox(frame_sort_order_column)
sort_order_column_combobox.pack(side='left', fill='x', expand=True)
sort_order_column_combobox.bind("<<ComboboxSelected>>", lambda event: update_sort_order_list())

# Sort Order Text Section (left column)
sort_order_frame = ttk.Frame(left_column_frame, padding="10")
sort_order_frame.pack(fill='x')

sort_order_label = ttk.Label(sort_order_frame, text="Sort Order:")
sort_order_label.pack(side='left')

sort_order_text = tk.Text(sort_order_frame, height=4, width=30)
sort_order_text.pack(side='left', padx=10)

# Group Size Entry (left column)
frame_group_size = ttk.Frame(left_column_frame, padding="10")
frame_group_size.pack(fill='x')

group_size_label = ttk.Label(frame_group_size, text="Group Size:")
group_size_label.pack(side='left')

group_size_entry = ttk.Entry(frame_group_size, textvariable=group_size_entry)
group_size_entry.pack(side='left')

# Right Column
right_column_frame = ttk.Frame(root, padding="10")
right_column_frame.grid(row=1, column=1, sticky='nsew', padx=5)

# Name Column Selection Section (right column)
frame_name_column = ttk.Frame(right_column_frame, padding="10")
frame_name_column.pack(fill='x')

name_column_label = ttk.Label(frame_name_column, text="Name Column:")
name_column_label.pack(side='left')

name_column_combobox = ttk.Combobox(frame_name_column, textvariable=name_column)
name_column_combobox.pack(side='left', fill='x', expand=True)

# Additional Columns Section (right column)
additional_columns_frame = ttk.Frame(right_column_frame, padding="10")
additional_columns_frame.pack(fill='both', expand=True)

# Add Additional Column Button (right column)
add_column_button = ttk.Button(additional_columns_frame, text="Extra Column", command=add_additional_column_selector)
add_column_button.pack(side='top', pady=10)

# Initialize with one additional column selector
add_additional_column_selector()

# Process File Button (spanning both columns)
process_button = ttk.Button(root, text="Process", command=process_file)
process_button.grid(row=2, column=0, columnspan=2, pady=10)

root.mainloop()
