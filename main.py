import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import load_workbook, Workbook
#COLORES
GRIS = "#d8cccc"
NEGRO = "#111111"

def Cargar_datos_Excel(archivo):# carga datos del excel
    wb = load_workbook(archivo)
    ws = wb.active
    registros = []
    for row in ws.iter_rows(min_row=2):
        if len(row) == 5:
            distribuidora_Name , cargador, distco, volumen, fecha = row
            registros.append((distribuidora_Name, cargador, distco, volumen, fecha))
        else:
            messagebox.showerror("Error en el registro de datos")
    
    return registros

    
def Cargar_Datos():#seleccionar el archivo a cargar
    archivo = filedialog.askopenfilename(title="Seleccionar archivo de Excel", filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")))
    if archivo:
        global registros
        registros = Cargar_datos_Excel(archivo)
        messagebox.showinfo("Éxito", "Los datos fueron cargados")

def Main_Menu():#funciones del menu principal
    for widget in root.winfo_children(): #limpiar la screen cada vez que llama la funcion
        widget.destroy()
    #Botones del meno principal
    btn_Cargar = tk.Button(root, text="Cargar Datos", bg=GRIS,fg=NEGRO,width=20,height=3, command=Cargar_Datos)
    btn_Cargar.pack(side=tk.LEFT, padx=10,pady=10)

    btn__Ver_Registros = tk.Button(root,text="Ver Registros", bg=GRIS,fg=NEGRO,width=20,height=3)
    btn__Ver_Registros.pack(side=tk.LEFT, padx=10, pady=10)

    btn_Calcular_Promedio = tk.Button(root,text="Calcular Promedio", bg=GRIS,fg=NEGRO,width=20,height=3)
    btn_Calcular_Promedio.pack(side=tk.LEFT,padx=10,pady=10)
root = tk.Tk()
root.title("Distribuidoras Mabel")
root.geometry("720x460")
root.config(bg="GRAY")
registros = []#para almacenar los registros
Main_Menu()
root.mainloop()