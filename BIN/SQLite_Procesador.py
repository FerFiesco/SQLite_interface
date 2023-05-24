import db_SQLliteV1_1 as db
import pandas as pd
from tkinter import filedialog
from tkinter import messagebox as msg
import os
import tkinter as tk
from tkinter import scrolledtext
from subprocess import run

run_path=os.getcwd()
out_path=os.path.join(str(run_path),"OUTPUT")
in_path=os.path.join(run_path, "INPUT")
config_path=os.path.join(run_path, "CONFIG")
bin_path=os.path.join(run_path, "BIN")

#Create folders if not exist
if not os.path.exists(out_path):    os.makedirs(out_path)
if not os.path.exists(in_path):    os.makedirs(in_path)
if not os.path.exists(config_path):    os.makedirs(config_path)
if not os.path.exists(bin_path):    os.makedirs(bin_path)

def run_full_process():
    update_data()
    create_tables()
    export_files()

# Funciones para cada botón
def update_data():
    print("Actualizando datos")
    select_input_files()
    print("Datos actualizados")

def create_tables():
    print("Creando tablas")
    #look for process file in config folder
    for file in os.listdir(config_path):
        if file.startswith("Process"):
            process_file  = os.path.join(config_path,file)
    with(open(process_file)) as f:
        query=f.read()
    for statement in query.split(';'):
        if statement.strip():
            print(statement)
            with db.engine.connect() as con:
                con.exec_driver_sql(statement)
    print("Tablas creadas")

def export_files():
    print("Exportando archivos")
    #look for export file in config folder
    for file in os.listdir(config_path):
        if file.startswith("Export"):
            export_file  = os.path.join(config_path,file) 

    with(open(export_file)) as f:
        file_text=f.read()
        db.create_file(file_text)
    print("Archivos exportados")

# Función para actualizar la caja de texto con la salida de la consola
def update_console(msg):
    console_output.configure(state='normal') # Desbloquear el estado de escritura de la caja de texto
    for m in msg:
        console_output.insert(tk.END, m + '\n') # Agregar mensaje a la caja de texto
    console_output.configure(state='disabled') # Bloquear el estado de escritura de la caja de texto
# Función para limpiar la caja de texto
def clear_console():
    console_output.configure(state='normal') # Desbloquear el estado de escritura de la caja de texto
    console_output.delete('1.0', tk.END) # Borrar todo el contenido de la caja de texto
    console_output.configure(state='disabled') # Bloquear el estado de escritura de la caja de texto

def select_input_files():
    file_path = filedialog.askopenfilename(title="Select input Files", filetypes=[("all files", "*.xlsx; *.xls ; *.csv")],initialdir=in_path,multiple=True)
    #file_dic = {}

    inputs_table = pd.read_csv(os.path.join(config_path,"inputs.csv"),)
    if len(file_path) != len(inputs_table.index):
        print("please select the correct number of files")
        msg.showinfo("Warning", "Only selected files goin to be uploaded")
        #exit()
    
    for file in file_path:
        updated=False
        for i in inputs_table.index:
            if str(inputs_table["Tag"][i]).upper() in str(os.path.basename(file)).upper():
                db.load_data_from_files(table= inputs_table["Table_name"][i], file_path= file) #file_dic["BOM"] = file
                updated=True
                break
        
        if not updated:
            print("file : " + os.path.basename(file) + " not recognized")
            msg.showwarning("Error", "file : " + os.path.basename(file) + " not recognized")
            pass
        #return file_dic


if __name__ == "__main__":
    print("starting")
    # Crear ventana
    window = tk.Tk()
    window.title("SQLite merge tool")

    # Definir colores y bordes
    bg_color = "#F5F5F5"
    button_bg_color = "#3F51B5"
    button_fg_color = "#FFFFFF"
    button_border_radius = 8

    # Crear frame para los botones
    button_frame = tk.Frame(window, bg=bg_color, padx=10, pady=10)

    # Crear botones y agregarlos al frame
    full_run_button = tk.Button(button_frame, text="Run all", command=run_full_process, padx=10, pady=10, anchor='nw', bg=button_bg_color, fg=button_fg_color, relief=tk.RIDGE)#, bd=0, borderwidth=0, highlightthickness=0, border=0, highlightcolor=button_bg_color, highlightbackground=button_bg_color, activebackground=button_bg_color, activeforeground=button_fg_color, cursor="hand2")
    update_button = tk.Button(button_frame, text="1-Update data", command=update_data, padx=10, pady=10, anchor='w', bg=button_bg_color, fg=button_fg_color, relief=tk.SOLID)#, bd=0, borderwidth=0, highlightthickness=0, border=0, highlightcolor=button_bg_color, highlightbackground=button_bg_color, activebackground=button_bg_color, activeforeground=button_fg_color, cursor="hand2")
    create_button = tk.Button(button_frame, text="2-Create tables", command=create_tables, padx=10, pady=10, anchor='w', bg=button_bg_color, fg=button_fg_color, relief=tk.SUNKEN)#, bd=0, borderwidth=0, highlightthickness=0, border=0, highlightcolor=button_bg_color, highlightbackground=button_bg_color, activebackground=button_bg_color, activeforeground=button_fg_color, cursor="hand2")
    export_button = tk.Button(button_frame, text="3-Export Files", command=export_files, padx=10, pady=10, anchor='w', bg=button_bg_color, fg=button_fg_color, relief=tk.FLAT)#, bd=0, borderwidth=0, highlightthickness=0, border=0, highlightcolor=button_bg_color, highlightbackground=button_bg_color, activebackground=button_bg_color, activeforeground=button_fg_color, cursor="hand2")

    full_run_button.pack(side=tk.TOP, pady=10)
    update_button.pack(side=tk.LEFT, padx=(0,10))
    create_button.pack(side=tk.LEFT, padx=(0,10))
    export_button.pack(side=tk.LEFT)

    # Agregar frame ala ventana
    button_frame.pack()

    # Crear cuadro de texto para imprimir la salida de la consola
    text_frame = tk.Frame(window, padx=10, pady=10)
    console_output = scrolledtext.ScrolledText(text_frame, height=10)
    #scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL)

    console_output.pack(fill=tk.BOTH, expand=True)
    #text_frame.pack(fill=tk.BOTH, expand=True)
    # Iniciar la aplicación
    window.mainloop()
        