import tkinter as tk
from tkinter import filedialog
import pandas as pd

icon_path = ".\excel.ico"

# Función para seleccionar un archivo de Excel
def select_excel_file(entry):
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(tk.END, filepath)

# Función para mostrar las hojas de cálculo del archivo de Excel seleccionado
def show_sheets(filepath, listbox):
    try:
        sheets = pd.ExcelFile(filepath).sheet_names
        listbox.delete(0, tk.END)
        for sheet in sheets:
            listbox.insert(tk.END, sheet)
    except Exception as e:
        listbox.delete(0, tk.END)
        listbox.insert(tk.END, "Error: " + str(e))

# Función para mostrar las columnas de la hoja de cálculo seleccionada
def show_columns(filepath, sheet, listbox):
    try:
        df = pd.read_excel(filepath, sheet_name=sheet)
        listbox.delete(0, tk.END)
        for column in df.columns:
            listbox.insert(tk.END, column)
    except Exception as e:
        listbox.delete(0, tk.END)
        listbox.insert(tk.END, "Error: " + str(e))

# Función para guardar la hoja de cálculo seleccionada
def save_selected_sheet(listbox, entry):
    selected_sheet = listbox.get(listbox.curselection())
    entry.delete(0, tk.END)
    entry.insert(tk.END, selected_sheet)

# Función para guardar la columna seleccionada
def save_selected_column(listbox, entry):
    selected_column = listbox.get(listbox.curselection())
    entry.delete(0, tk.END)
    entry.insert(tk.END, selected_column)

# Función para realizar el merge con la columna seleccionada
def merge_excel_files():
    file1 = excel_entry1.get()
    file2 = excel_entry2.get()
    selected_sheet = selected_sheet_entry.get()
    selected_column = selected_column_entry.get()

    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2, sheet_name=selected_sheet)

        # Convertir la columna del segundo archivo a tipo objeto
        df2[selected_column] = df2[selected_column].astype(object)

        # Realizar el merge utilizando la columna seleccionada
        merged_df = pd.merge(df1, df2, how='left', on=selected_column)

        # Generar un nuevo archivo para el merge
        merge_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        with pd.ExcelWriter(merge_file_path, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Merged')

        result_label.config(text="Merge completado. Los datos fusionados se han guardado en el archivo merge_file.xlsx.")
    except Exception as e:
        result_label.config(text="Error en el merge: " + str(e))

# Crear la ventana principal
window = tk.Tk()
window.title("Merge Excel Files")
window.geometry("700x800")
window.iconbitmap(icon_path)

# Primer archivo de Excel
excel_label1 = tk.Label(window, text="Primer archivo de Excel:")
excel_label1.pack()

excel_entry1 = tk.Entry(window, width=40)
excel_entry1.pack()

excel_button1 = tk.Button(window, text="Seleccionar archivo", command=lambda: select_excel_file(excel_entry1))
excel_button1.pack()

# Segundo archivo de Excel
excel_label2 = tk.Label(window, text="Segundo archivo de Excel:")
excel_label2.pack()

excel_entry2 = tk.Entry(window, width=40)
excel_entry2.pack()

excel_button2 = tk.Button(window, text="Seleccionar archivo", command=lambda: select_excel_file(excel_entry2))
excel_button2.pack()

# Hojas de cálculo del segundo archivo de Excel
show_sheets_label = tk.Label(window, text="Hojas de cálculo del segundo archivo:")
show_sheets_label.pack()

sheet_listbox = tk.Listbox(window)
sheet_listbox.pack()

show_sheets_button = tk.Button(window, text="Mostrar hojas de cálculo", command=lambda: show_sheets(excel_entry2.get(), sheet_listbox))
show_sheets_button.pack()

# Hoja de cálculo seleccionada
selected_sheet_label = tk.Label(window, text="Seleccionar hoja de cálculo:")
selected_sheet_label.pack()

selected_sheet_entry = tk.Entry(window, width=40)
selected_sheet_entry.pack()

save_sheet_button = tk.Button(window, text="Guardar hoja de cálculo seleccionada", command=lambda: save_selected_sheet(sheet_listbox, selected_sheet_entry))
save_sheet_button.pack()

# Columnas de la hoja de cálculo seleccionada
show_columns_label = tk.Label(window, text="Columnas de la hoja de cálculo seleccionada:")
show_columns_label.pack()

column_listbox = tk.Listbox(window)
column_listbox.pack()

show_columns_button = tk.Button(window, text="Mostrar columnas", command=lambda: show_columns(excel_entry2.get(), selected_sheet_entry.get(), column_listbox))
show_columns_button.pack()

# Columna seleccionada
selected_column_label = tk.Label(window, text="Seleccionar columna:")
selected_column_label.pack()

selected_column_entry = tk.Entry(window, width=40)
selected_column_entry.pack()

save_column_button = tk.Button(window, text="Guardar columna seleccionada", command=lambda: save_selected_column(column_listbox, selected_column_entry))
save_column_button.pack()

# Realizar el merge
merge_button = tk.Button(window, text="Realizar Merge", command=merge_excel_files)
merge_button.pack()

# Etiqueta para mostrar el resultado del merge
result_label = tk.Label(window, text="")
result_label.pack()

# Ejecutar el bucle principal de la ventana
window.mainloop()
