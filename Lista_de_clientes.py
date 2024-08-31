import customtkinter as ctk
from tkinter import *
from tkinter import messagebox, filedialog
import openpyxl
import pathlib
from openpyxl import Workbook
import os

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.appearance()
        self.selected_dir = "dados_clientes"
        self.create_dir(self.selected_dir)
        self.todo_sistema()

    def layout_config(self):  
        self.title("Sistema de Gestão de Clientes")
        self.geometry("700x500")

    def appearance(self):     
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="transparent", text_color=['#000', "#fff"]).place(x=50, y=430)
        self.opt_apm = ctk.CTkOptionMenu(self, values=["Light", "Dark", "System"], command=self.change_apm).place(x=50, y=460)
    
    def create_dir(self, dir_name):
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
        
    def select_directory(self):
        dir_selected = filedialog.askdirectory()
        if dir_selected:
            self.selected_dir = dir_selected
            self.create_dir(self.selected_dir)
            messagebox.showinfo("Sistema", f"Diretório selecionado: {self.selected_dir}")
    
    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=700, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)
        title = ctk.CTkLabel(frame, text="Sistema de Gestão de Clientes", font=("Century Gothic bold", 24), text_color="#fff").place(x=180, y=10)
        span = ctk.CTkLabel(self, text="Por favor, preencha todos os campos do formulário", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=50, y=70)

        def submit():
            # Os dados dos entrys 
            name = name_value.get()
            contact = contact_value.get()
            age = age_value.get()
            address = address_value.get()
            gender = gender_combobox.get()
            obs = obs_entry.get("1.0", END)

            if name == "" or contact == "" or age == "" or address == "":
                messagebox.showerror("Sistema", "Erro!\nPor favor, preencha todos os campos!")
            else:
                file_path = os.path.join(self.selected_dir, "Clientes.xlsx")
                ficheiro = openpyxl.load_workbook(file_path) if os.path.exists(file_path) else Workbook()
                folha = ficheiro.active
                if not os.path.exists(file_path):
                    folha['A1'] = "Nome completo"
                    folha['B1'] = "Contato"
                    folha['C1'] = "Idade"
                    folha['D1'] = "Gênero"
                    folha['E1'] = "Endereço"
                    folha['F1'] = "Observações"
                folha.append([name, contact, age, gender, address, obs.strip()])
                ficheiro.save(file_path)
                messagebox.showinfo("Sistema", "Dados salvos com sucesso!")
        
        def clear():
            name_value.set("")
            contact_value.set("")
            age_value.set("")
            address_value.set("")
            obs_entry.delete("1.0", END)
        
        # Variables 
        name_value = StringVar()
        contact_value = StringVar()
        age_value = StringVar()
        address_value = StringVar()
        
        # Entrys
        name_entry = ctk.CTkEntry(self, width=250, textvariable=name_value, font=("Century Gothic bold", 16), fg_color="transparent")
        contact_entry = ctk.CTkEntry(self, width=200, textvariable=contact_value, font=("Century Gothic bold", 16), fg_color="transparent")
        age_entry = ctk.CTkEntry(self, width=100, textvariable=age_value, font=("Century Gothic bold", 16), fg_color="transparent")
        address_entry = ctk.CTkEntry(self, width=250, textvariable=address_value, font=("Century Gothic bold", 16), fg_color="transparent")
    
        # Combobox
        gender_combobox = ctk.CTkComboBox(self, values=["Masculino", "Feminino", "Outros..."], font=("Century Gothic bold", 14))
        gender_combobox.set("......")
    
        # Entrada de obs
        obs_entry = ctk.CTkTextbox(self, width=500, height=120, font=("arial", 18), border_color="#aaa", border_width=2, fg_color="transparent")
    
        # Labels
        lb_name = ctk.CTkLabel(self, text="Nome", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=50, y=120)
        lb_contact = ctk.CTkLabel(self, text="Contato", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=350, y=120)
        lb_age = ctk.CTkLabel(self, text="Idade", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=50, y=190)
        lb_gender = ctk.CTkLabel(self, text="Gênero", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=200, y=190)
        lb_address = ctk.CTkLabel(self, text="Endereço", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=350, y=190)
        lb_obs = ctk.CTkLabel(self, text="Observações", font=("Century Gothic bold", 16), text_color=["#000", "#fff"]).place(x=50, y=260)
    
        btn_submit = ctk.CTkButton(self, text="Salvar dados".upper(), command=submit, fg_color="#151", hover_color="#131").place(x=300, y=420)
        btn_clear = ctk.CTkButton(self, text="Limpar campos".upper(), command=clear, fg_color="#555", hover_color="#333").place(x=500, y=420)
        btn_select_dir = ctk.CTkButton(self, text="Selecionar diretório", command=self.select_directory, fg_color="#151", hover_color="#131").place(x=50, y=420)
        
        # Posicionamento
        name_entry.place(x=50, y=150)
        contact_entry.place(x=350, y=150)
        age_entry.place(x=50, y=220)
        gender_combobox.place(x=200, y=220)
        address_entry.place(x=350, y=220)
        obs_entry.place(x=50, y=290)

    def change_apm(self, new_appearance):
        ctk.set_appearance_mode(new_appearance)

if __name__ == "__main__":
    app = App()
    app.mainloop()
