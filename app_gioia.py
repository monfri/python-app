#librerie tkinter e custom tk per creazione finestre
import customtkinter as ctk
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import os
import pandas as pd

def select_file(label,imput_dialog_title,file_format):
    file = filedialog.askopenfilenames(
        initialdir=os.getcwd(),
        title=imput_dialog_title,
            filetypes=file_format)
    if file:
#        label.configure(text=os.path.basename(file[0]))
        label.configure(text=file[0])
    
def elimina_doppioni():
    global file_cfg
    #Abilitazione della combobox projects
    if excel_label.cget('text')!='percorso file':
        file_cfg=excel_label.cget('text') 
        df = pd.read_excel(file_cfg)
        
        # Convertire tutte le email in minuscolo
        df['EMAIL CLIENTE'] = df['EMAIL CLIENTE'].str.lower()
            
        # Identificare le email duplicate
        duplicates = df[df.duplicated(subset='EMAIL CLIENTE', keep=False)]

        # Rimuovere le righe con email duplicate
        df_cleaned = df[~df['EMAIL CLIENTE'].isin(duplicates['EMAIL CLIENTE'])]
        
        # Ordinare i duplicati in ordine alfabetico in base alla colonna 'EMAIL CLIENTE'
        duplicates_sorted = duplicates.sort_values(by='EMAIL CLIENTE')

        # Percorso del nuovo file Excel
        output_path = file_cfg[:-5]+"_Elaborato.xlsx"

        # Scrivere i DataFrame su due fogli diversi nel nuovo file Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_cleaned.to_excel(writer, sheet_name='Senza Duplicati', index=False)
            duplicates_sorted.to_excel(writer, sheet_name='Duplicati Rimossi', index=False)
        start_label.configure(text="Doppioni rimossi")

####SCRIPT PRINCIPALE: CREAZIONE FINESTRA MAIN E WIDGET
if __name__ == "__main__":    
    # Creare la finestra principale
    ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
    ctk.set_default_color_theme("green")  # Themes: blue (default), dark-blue, green
    main = ctk.CTk()
    main.title("ISA data analysis")
    width =600
    heigth =400
    x = (main.winfo_screenwidth()/2) - (width/2)
    y = (main.winfo_screenheight()/2) - (heigth/2)
    main.geometry(f'{width}x{heigth}+{int(x)}+{int(y)}')
    main.grid_columnconfigure(0, weight=1)
    main.grid_rowconfigure(0, weight=1)
    
    # FRAME PRINCIPALE
    main_frame = ctk.CTkFrame(main)
    main_frame.grid(row=0, column=0, padx=10, pady=(10, 10), sticky="nesw")
    main_frame.columnconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=1)
    main_frame.columnconfigure(2, weight=1)
    main_frame.columnconfigure(3, weight=1)
    
    excel_button = ctk.CTkButton(main_frame,text="Carica file excel",  compound="left", height= 50, width=580, command = lambda:  select_file(excel_label,"Selezionare il file excel",(("excel files","*.xlsx"),("All files","*.*"))))
    excel_button.grid(row=0, column=0, padx=(10,10), pady=20, sticky="ew")
    
    excel_label = ctk.CTkLabel(main_frame,text='percorso file', width= 200,height=70,fg_color="#2A2A2A",corner_radius = 20,wraplength=200)
    excel_label.grid(row=1, column=0, padx=(10,10), pady=(5,20), sticky="ew")
    
    start_button = ctk.CTkButton(main_frame,text="Start",  compound="left", height= 50, width=580, command = elimina_doppioni)
    start_button.grid(row=2, column=0, padx=(10,10), pady=20, sticky="ew")
    
    start_label = ctk.CTkLabel(main_frame,text='', width= 200,height=70,fg_color="#2A2A2A",corner_radius = 20,wraplength=200)
    start_label.grid(row=3, column=0, padx=(10,10), pady=(5,20), sticky="ew")
    
    
    
    
    
    
    
    main.mainloop()