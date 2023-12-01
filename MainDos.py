import os
import pandas as pd
from tkinter import Tk, Button, Label, filedialog, Text, Scrollbar, Frame, RIGHT, Y, END, messagebox

class CSVAnalyzer:
    def __init__(self, master):
        self.master = master
        master.title("CSV Analyzer")

        self.files = []
        self.dfs = []
        self.combined_df = None  # Agregamos combined_df como atributo

        self.label = Label(master, text="Seleccione hasta 3 archivos CSV:")
        self.label.pack()

        self.choose_button = Button(master, text="Seleccionar archivos", command=self.choose_files)
        self.choose_button.pack()

        self.process_button = Button(master, text="Procesar archivos", command=self.process_files, state="disabled")
        self.process_button.pack()

        self.text_frame = Frame(master)
        self.text_frame.pack()

        self.text_widget = Text(self.text_frame, wrap="none", width=80, height=20)
        self.text_widget.pack(side="left", fill="y")

        self.scrollbar = Scrollbar(self.text_frame, command=self.text_widget.yview)
        self.scrollbar.pack(side="right", fill="y")

        self.text_widget.config(yscrollcommand=self.scrollbar.set)

        # Nuevo botón para guardar el DataFrame combinado en un archivo Excel
        self.save_button = Button(master, text="Guardar", command=self.save_to_excel, state="disabled")
        self.save_button.pack()

    def choose_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Seleccione archivos CSV",
            filetypes=(("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")),
            initialdir="/"
        )

        if file_paths:
            self.files = file_paths[:3]  # Tomar hasta 3 archivos
            self.process_button["state"] = "normal"  # Habilitar el botón de procesar

    def process_files(self):
        unique_barcodes = set()  # Conjunto para almacenar códigos de barras únicos

        # Iterar sobre los archivos y extraer códigos de barras únicos
        for file_path in self.files:
            df = pd.read_csv(file_path, header=None, encoding='latin1', sep=';')  # Especificar el separador

            # Seleccionar solo las columnas especificadas
            df = df[[1, 5, 9]]

            # Renombrar las columnas
            df = df.rename(columns={1: 'CodBarra', 5: 'Nombre', 9: 'Precio'})

            # Agregar una columna 'NOMBREARCHIVO' con el nombre del archivo (sin la ruta)
            df['NOMBREARCHIVO'] = os.path.basename(file_path)

            # Eliminar duplicados basados en 'CodBarra' y mantener la primera ocurrencia
            df = df.drop_duplicates(subset=['CodBarra'])

            unique_barcodes.update(df['CodBarra'])  # Agregar códigos de barras únicos al conjunto
            self.dfs.append(df)

        # Crear un DataFrame nuevo con códigos de barras únicos como índice (convertir a lista)
        self.combined_df = pd.DataFrame(index=list(unique_barcodes))

        
        # Iterar sobre los DataFrames y actualizar el DataFrame combinado
        for df in self.dfs:
            # Actualizar el DataFrame combinado con los precios correspondientes
            self.combined_df[f'Nombre_{df["NOMBREARCHIVO"].iloc[0]}'] = df.set_index('CodBarra')['Nombre']
            self.combined_df[f'Precio_{df["NOMBREARCHIVO"].iloc[0]}'] = df.set_index('CodBarra')['Precio']
            
        # Habilitar el botón de guardar
        self.save_button["state"] = "normal"

        # Puedes imprimir el DataFrame combinado en la ventana emergente
        self.print_to_text_widget("\nDataFrame Combinado:\n")
        self.print_to_text_widget(self.combined_df)
        self.print_to_text_widget("\n")

    def save_to_excel(self):
        if self.combined_df is not None:
            # Permitir al usuario seleccionar la ubicación para guardar el Excel
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivo Excel", "*.xlsx")])
            if save_path:
                try:
                    # Convertir DataFrame combinado a Excel
                    
                    self.combined_df.reset_index().to_excel(save_path, index=False)  # Resetear el índice antes de guardar

                    self.print_to_text_widget(f"\nDataFrame Combinado guardado en: {save_path}\n")
                except Exception as e:
                    messagebox.showerror("Error", f"Error al guardar en Excel: {str(e)}")
        else:
            messagebox.showwarning("Advertencia", "No hay datos para guardar en Excel.")

    def print_to_text_widget(self, content):
        self.text_widget.insert(END, content)
        self.text_widget.see(END)  # Desplazar hasta el final

if __name__ == "__main__":
    root = Tk()
    app = CSVAnalyzer(root)
    root.mainloop()
