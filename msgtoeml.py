import os
import mimetypes
import sys
import webbrowser
import extract_msg
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from email.message import EmailMessage

class CTkDnDWindow(TkinterDnD.Tk, ctk.CTkBaseClass):
    def __init__(self, *args, **kwargs):
        TkinterDnD.Tk.__init__(self, *args, **kwargs)
        ctk.CTkBaseClass.__init__(self, *args, **kwargs)
# Główna aplikacja
class App(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()


        self.title("MSG  ➜ EML Konwerter z załącznikami ver. 1.0 autor: Ernest Zając")
        self.geometry("900x450")
        self.resizable(False,False)

        # Konfiguracja stylu
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        path = os.path.join(os.path.dirname(sys.modules['customtkinter'].__file__), "assets", "icons",
                            "customtkinter_icon_Windows.ico")
        self.iconbitmap(path)


        # Ramka główna
        self.ctk_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.ctk_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.msg_path = None
        self.eml_path = None

        # Nagłówek
        self.label = ctk.CTkLabel(self.ctk_frame, text="Konwerter (wiadomości Outlook) .MSG ➜ .EML (wiadomości Thunderbird)", font=("Arial", 20, "bold"))
        self.label.pack(pady=10)

        # Przycisk wyboru pliku
        self.select_button = ctk.CTkButton(self.ctk_frame, text="Wybierz plik MSG", command=self.select_file)
        self.select_button.pack(pady=10)

        # Informacja o pliku
        self.file_label = ctk.CTkLabel(self.ctk_frame, text="Nie wybrano pliku")
        self.file_label.pack()

        # Pole do przeciągania pliku
        self.drop_label = ctk.CTkLabel(self.ctk_frame, text="Przeciągnij plik MSG tutaj", fg_color="#dddddd", height=80, corner_radius=10)
        self.drop_label.pack(padx=20, pady=10, fill="x")
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.drop_file)

        # Pasek postępu
        self.progress = ctk.CTkProgressBar(self.ctk_frame)
        self.progress.set(0)
        self.progress.pack(padx=20, pady=10, fill="x")

        # Ramka przycisków
        self.buttons_frame = ctk.CTkFrame(self.ctk_frame, fg_color="transparent")
        self.buttons_frame.pack(pady=20)

        self.clear_button = ctk.CTkButton(self.buttons_frame, text="Wyczyść", command=self.clear)
        self.clear_button.grid(row=0, column=0, padx=10)

        self.save_button = ctk.CTkButton(self.buttons_frame, text="Zapisz jako EML", command=self.convert_and_save)
        self.save_button.grid(row=0, column=1, padx=10)

        self.open_button = ctk.CTkButton(self.buttons_frame, text="Otwórz EML", command=self.open_eml)
        self.open_button.grid(row=0, column=2, padx=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("MSG Files", "*.msg")])
        if file_path:
            self.set_file(file_path)

    def drop_file(self, event):
        file_path = event.data.strip("{}")
        if file_path.lower().endswith(".msg"):
            self.set_file(file_path)
        else:
            messagebox.showerror("Błąd", "Obsługiwany jest tylko format .msg")

    def set_file(self, path):
        self.msg_path = path
        self.file_label.configure(text=f"Wybrano: {os.path.basename(path)}")
        self.progress.set(0.5)

    def convert_and_save(self):
        if not self.msg_path:
            messagebox.showwarning("Brak pliku", "Wybierz lub przeciągnij plik MSG.")
            return

        try:
            msg = extract_msg.Message(self.msg_path)
            msg_subject = msg.subject or "Brak tematu"
            msg_sender = msg.sender or "Brak nadawcy"
            msg_to = msg.to or "Brak odbiorcy"
            msg_date = msg.date or "Brak daty"
            msg_body = msg.body or "(pusta wiadomość)"

            email_msg = EmailMessage()
            email_msg["Subject"] = msg_subject
            email_msg["From"] = msg_sender
            email_msg["To"] = msg_to
            email_msg["Date"] = msg_date
            email_msg.set_content(msg_body)

            # Dodaj załączniki
            for attachment in msg.attachments:
                filename = attachment.longFilename or attachment.shortFilename or "zalacznik"
                data = attachment.data
                mime_type, _ = mimetypes.guess_type(filename)
                maintype, subtype = mime_type.split("/") if mime_type else ("application", "octet-stream")

                email_msg.add_attachment(
                    data,
                    maintype=maintype,
                    subtype=subtype,
                    filename=filename
                )

            # Wybór miejsca zapisu
            default_name = os.path.splitext(os.path.basename(self.msg_path))[0] + ".eml"
            save_path = filedialog.asksaveasfilename(defaultextension=".eml", initialfile=default_name)

            if save_path:
                with open(save_path, "wb") as f:
                    f.write(email_msg.as_bytes())
                self.eml_path = save_path
                self.progress.set(1)
                messagebox.showinfo("Sukces", "Plik EML zapisany z załącznikami.")

        except Exception as e:
            messagebox.showerror("Błąd konwersji", str(e))

    def open_eml(self):
        if self.eml_path and os.path.exists(self.eml_path):
            webbrowser.open(self.eml_path)
        else:
            messagebox.showwarning("Brak pliku", "Najpierw zapisz plik EML.")

    def clear(self):
        self.msg_path = None
        self.eml_path = None
        self.file_label.configure(text="Nie wybrano pliku")
        self.progress.set(0)

if __name__ == "__main__":
    app = App()
    app.mainloop()
