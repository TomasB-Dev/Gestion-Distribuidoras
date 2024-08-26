import tkinter as tk
import openpyxl
#COLORES
GRIS = "#d8cccc"
NEGRO = "#111111"

def Main_Menu():#funciones del menu principal
    for widget in root.winfo_children(): #limpiar la screen cada vez que llama la funcion
        widget.destroy()
    #Botones del meno principal
    btn_Cargar = tk.Button(root, text="Cargar Datos", bg=GRIS,fg=NEGRO,width=20,height=3)
    btn_Cargar.pack(side=tk.LEFT, padx=10,pady=10)
root = tk.Tk()
root.title("Distribuidoras Mabel")
root.geometry("720x460")
root.config(bg="GRAY")

Main_Menu()
root.mainloop()