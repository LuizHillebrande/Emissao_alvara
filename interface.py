import customtkinter as ctk


ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  


app = ctk.CTk()
app.title("Office automation")
app.geometry("400x300")  


label = ctk.CTkLabel(app, text="Controle de débitos municipais!", font=("Arial", 16))
label.pack(pady=20)

button = ctk.CTkButton(app, text="Clique Aqui", command=lambda: print("Botão clicado!"))
button.pack(pady=10)

app.mainloop()
