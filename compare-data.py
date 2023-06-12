# # # importing the module

# # import pandas as pd

# # firmas = pd. read_excel('firmas_licenciados.xlsx')
# # personal = pd.read_excel("personal.xlsx")
# # result = pd.merge(firmas, personal,on='Nombre_Personal', how='left')

# # print(result.head())
# # result.to_excel("Results.xlsx", index = False)

# # importing the module
# import pandas

# # reading the files
# f1 = pandas.read_excel("Medicos_Detalle.xlsx")
# f2 = pandas.read_excel(r"E:\HCE - IMPORTANTES\USUARIOS HOSIX ACTIVOS.xlsx", sheet_name="MÉDICOS")

# f1['DNI'] = f1['DNI'].astype(str)
# f2['DNI'] = f2['DNI'].astype(str)

# print(f2.columns.tolist())

# # merging the files
# f3 = f1[["DNI",
#          "Total_Firma_Pendiente", "Nombre_Medico"]].merge(f2[["DNI",
#                                                               "Numero"]],
#                                                           on="DNI",
#                                                           how="left")

# # creating a new file
# f3.to_excel("Resultado-25-05-23.xlsx", index=False)
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def show_common_columns():
    file1_path = file1_var.get()
    file2_path = file2_var.get()

    if file1_path and file2_path:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        common_columns = set(df1.columns) & set(df2.columns)

        # Mostrar las columnas comunes en el cuadro de texto
        columns_text.delete(1.0, tk.END)
        columns_text.insert(tk.END, ', '.join(common_columns))
    else:
        columns_text.delete(1.0, tk.END)
        columns_text.insert(tk.END, 'Seleccione ambos archivos para mostrar las columnas comunes.')

def merge_files():
    file1_path = file1_var.get()
    file2_path = file2_var.get()
    selected_column = column_var.get()

    if file1_path and file2_path and selected_column:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)

        if selected_column in df1.columns and selected_column in df2.columns:
            merged_df = pd.merge(df1, df2, on=selected_column, how='inner')

            # Guardar la información fusionada en el primer archivo
            merged_df.to_excel(file1_path, index=False)

            print("Merge completado y datos guardados en el archivo seleccionado.")
        else:
            print("La columna seleccionada no es común en ambos archivos.")
    else:
        print('Por favor, selecciona ambos archivos y una columna en común.')

def select_file1():
    file_path = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    file1_var.set(file_path)

def select_file2():
    file_path = filedialog.askopenfilename(filetypes=[('Archivos Excel', '*.xlsx')])
    file2_var.set(file_path)

root = tk.Tk()
root.title("Merge de Archivos Excel")

# Variables de control
file1_var = tk.StringVar()
file2_var = tk.StringVar()
column_var = tk.StringVar()

# Frame para el primer archivo
file1_frame = tk.Frame(root)
file1_frame.pack()

file1_label = tk.Label(file1_frame, text="Archivo 1:")
file1_label.pack(side=tk.LEFT)

file1_entry = tk.Entry(file1_frame, textvariable=file1_var)
file1_entry.pack(side=tk.LEFT)

file1_button = tk.Button(file1_frame, text="Seleccionar", command=select_file1)
file1_button.pack(side=tk.LEFT)

# Frame para el segundo archivo
file2_frame = tk.Frame(root)
file2_frame.pack()

file2_label = tk.Label(file2_frame, text="Archivo 2:")
file2_label.pack(side=tk.LEFT)

file2_entry = tk.Entry(file2_frame, textvariable=file2_var)
file2_entry.pack(side=tk.LEFT)

file2_button = tk.Button(file2_frame, text="Seleccionar", command=select_file2)
file2_button.pack(side=tk.LEFT)

# Botón para mostrar las columnas comunes
columns_button = tk.Button(root, text="Mostrar Columnas Comunes", command=show_common_columns)
columns_button.pack()

# Cuadro de texto para mostrar las columnas comunes
columns_text = tk.Text(root, height=5, width=50)
columns_text.pack()

# Frame para la columna en común
column_frame = tk.Frame(root)
column_frame.pack()

column_label = tk.Label(column_frame, text="Columna en común:")
column_label.pack(side=tk.LEFT)

column_entry = tk.Entry(column_frame, textvariable=column_var)
column_entry.pack(side=tk.LEFT)

# Botón para hacer el merge
merge_button = tk.Button(root, text="Merge y Guardar", command=merge_files)
merge_button.pack()

root.mainloop()
