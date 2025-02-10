import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import csv
import smtplib
import os  # Añadir esta línea
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class EmailSenderGUI:
    def __init__(self, master):
        self.master = master
        master.title("Sistema de Envío Masivo de Emails")
        master.geometry("900x800")
        
        # Establecer tema personalizado con colores marrones
        style = ttk.Style()
        style.configure(".", background="#f5e6d3")  # Fondo marrón muy claro
        style.configure("TLabel", background="#f5e6d3", foreground="#6b4423")  # Texto marrón oscuro
        style.configure("TButton", background="#8b7355", foreground="white")  # Botones marrones
        style.configure("TFrame", background="#f5e6d3")
        style.configure("TLabelframe", background="#f5e6d3")
        style.configure("TLabelframe.Label", background="#f5e6d3", foreground="#6b4423")

        # Crear container principal
        container = ttk.Frame(master)
        container.pack(fill="both", expand=True)
        
        # Configurar grid del container
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Crear canvas con scrollbar
        self.canvas = tk.Canvas(container)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        # Configurar el canvas
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Hacer que el frame scrollable ocupe todo el ancho disponible
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        
        # Crear ventana dentro del canvas que se expanda horizontalmente
        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame,
            anchor="nw",
            width=self.canvas.winfo_width()  # Establecer el ancho inicial
        )
        
        # Configurar el canvas para que se expanda
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Colocar canvas y scrollbar usando grid
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Vincular eventos de redimensionamiento
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
        # Configurar el scroll con el mouse
        self.scrollable_frame.bind('<Enter>', self._bound_to_mousewheel)
        self.scrollable_frame.bind('<Leave>', self._unbound_to_mousewheel)
        
        # Variables para configuración SMTP y email
        self.smtp_server_var = tk.StringVar()
        self.smtp_port_var = tk.StringVar(value="587")
        self.smtp_user_var = tk.StringVar()
        self.smtp_password_var = tk.StringVar()
        self.from_email_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        
        # Lista de destinatarios (se llenará al importar CSV o pegar emails)
        self.recipients = []  # Lista de diccionarios: {"email": ..., "nombre": ...}
        
        # --- Header con logo y título ---
        header = ttk.Frame(self.scrollable_frame)  # Quitamos el bootstyle="primary" para evitar el azul
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        header.configure(style="TFrame")  # Aplicamos el estilo marrón
        
        # Frame para el logo con fondo marrón
        logo_frame = ttk.Frame(header, style="TFrame")
        logo_frame.pack(pady=10)
        
        try:
            # Cargar y mostrar el logo
            logo_img = tk.PhotoImage(file="logo.png")
            # Redimensionar el logo a un tamaño más pequeño
            logo_img = logo_img.subsample(5, 5)  # Hacemos el logo aún más pequeño
            logo_label = ttk.Label(
                logo_frame, 
                image=logo_img, 
                background="#f5e6d3"  # Fondo marrón claro igual que el resto
            )
            logo_label.image = logo_img
            logo_label.pack(pady=(5, 5))
        except:
            ttk.Label(
                logo_frame,
                text="Logo no encontrado",
                font=("Helvetica", 10),
                background="#f5e6d3",
                foreground="#6b4423"
            ).pack(pady=5)

        # Título con fondo marrón
        ttk.Label(
            header,
            text="Sistema de Envío de Emails",
            font=("Helvetica", 16, "bold"),  # Añadimos "bold" para destacar
            background="#d4b08c",
            foreground="#4a3728",
            padding=10,
            style="TLabel"
        ).pack(fill="x")

        # --- Configuración SMTP ---
        smtp_frame = ttk.Labelframe(self.scrollable_frame, text="Configuración SMTP", padding=10)
        smtp_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        smtp_frame.columnconfigure(1, weight=1)
        
        labels = ["Servidor SMTP:", "Puerto:", "Usuario:", "Contraseña:", "Email remitente:"]
        vars = [self.smtp_server_var, self.smtp_port_var, self.smtp_user_var, 
                self.smtp_password_var, self.from_email_var]
        
        for idx, (label, var) in enumerate(zip(labels, vars)):
            ttk.Label(smtp_frame, text=label).grid(row=idx, column=0, sticky="w", padx=5, pady=2)
            if label == "Contraseña:":
                ttk.Entry(smtp_frame, textvariable=var, show="*").grid(
                    row=idx, column=1, sticky="ew", padx=5, pady=2)
            else:
                ttk.Entry(smtp_frame, textvariable=var).grid(
                    row=idx, column=1, sticky="ew", padx=5, pady=2)

        # --- Asunto del Email ---
        subject_frame = ttk.Labelframe(self.scrollable_frame, text="Asunto del Email", padding=10)
        subject_frame.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        ttk.Entry(subject_frame, textvariable=self.subject_var).pack(fill="x", padx=5, pady=2)

        # --- Lista de Destinatarios ---
        recipients_frame = ttk.Labelframe(
            self.scrollable_frame, 
            text="Lista de Emails (separados por comas o importar CSV)",
            padding=10
        )
        recipients_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")
        
        ttk.Button(
            recipients_frame,
            text="Importar CSV",
            command=self.import_csv,
            style="primary.TButton"  # Estilo personalizado para botones
        ).pack(side="top", anchor="ne", padx=5, pady=2)
        
        self.recipients_text = ScrolledText(recipients_frame, height=5)
        self.recipients_text.pack(fill="both", expand=True, padx=5, pady=2)

        # --- Cuerpo del Email ---
        message_frame = ttk.Labelframe(self.scrollable_frame, text="Cuerpo del Email", padding=10)
        message_frame.grid(row=4, column=0, padx=10, pady=5, sticky="nsew")
        self.message_text = ScrolledText(message_frame, height=10)
        self.message_text.pack(fill="both", expand=True, padx=5, pady=2)

        # --- Variables Disponibles (Tutorial) ---
        tutorial_frame = ttk.Labelframe(self.scrollable_frame, text="Variables Disponibles", padding=10)
        tutorial_frame.grid(row=5, column=0, padx=10, pady=5, sticky="ew")
        tutorial_msg = (
            "Puedes usar las siguientes variables en el cuerpo del email:\n"
            "  {email}  - Dirección de correo del destinatario\n"
            "  {nombre} - Nombre del destinatario (si está en el CSV; de lo contrario se usará el email)"
        )
        ttk.Label(tutorial_frame, text=tutorial_msg).pack(anchor="w")

        # --- Log de Envío ---
        log_frame = ttk.Labelframe(self.scrollable_frame, text="Log de Envío", padding=10)
        log_frame.grid(row=6, column=0, padx=10, pady=5, sticky="nsew")
        self.log_text = ScrolledText(log_frame, height=10)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=2)

        # --- Botón Enviar ---
        ttk.Button(
            self.scrollable_frame,
            text="Enviar Emails",
            command=self.send_emails,
            style="success.TButton",
            width=20
        ).grid(row=7, column=0, pady=10)

        # --- Footer ---
        footer = ttk.Frame(self.scrollable_frame)
        footer.grid(row=8, column=0, sticky="ew", pady=10)
        ttk.Label(
            footer,
            text="Diseñado con ❤️ por franHR\nCopyright © 2025 pcprogramacion.es",
            justify="center"
        ).pack()

        # Configuración de peso para que se redimensione bien la ventana
        self.scrollable_frame.grid_rowconfigure(4, weight=1)
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
    
    def _bound_to_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbound_to_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def import_csv(self):
        """Importa una lista de destinatarios desde un archivo CSV."""
        try:
            file_path = filedialog.askopenfilename(
                title="Seleccionar archivo CSV",
                filetypes=[("Archivos CSV", "*.csv")]
            )
            if not file_path:
                return
            
            self.recipients = []
            with open(file_path, 'r', encoding='utf-8') as f:
                # Usar siempre coma
                reader = csv.DictReader(f, delimiter=',')

                # Normalizar las columnas a minúsculas
                fieldnames_lower = [col.lower() for col in reader.fieldnames]
                if 'nombre' not in fieldnames_lower or 'email' not in fieldnames_lower:
                    messagebox.showerror(
                        "Error",
                        "El CSV debe contener columnas 'nombre' y 'email'."
                    )
                    return
                
                valid_count = 0
                for row in reader:
                    # Acceder usando las columnas normalizadas
                    nombre = row.get('nombre') or row.get('Nombre') or ''
                    email = row.get('email') or row.get('Email') or ''
                    nombre = nombre.strip()
                    email = email.strip()
                    if '@' in email:
                        self.recipients.append({"nombre": nombre or email, "email": email})
                        valid_count += 1
            
            if self.recipients:
                self.recipients_text.delete("1.0", tk.END)
                self.recipients_text.insert(tk.END, ", ".join([r["email"] for r in self.recipients]))
                self.log(f"Se importaron {valid_count} destinatarios.", "success")
                messagebox.showinfo("Éxito", f"CSV cargado: {valid_count} emails válidos.")
            else:
                messagebox.showerror("Error", "No se encontraron filas válidas (nombre,email).")

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo leer el CSV:\n{str(e)}\nAsegúrate de usar columnas 'nombre' y 'email'."
            )
            self.log(f"Error en import_csv: {e}", "error")
    
    def log(self, message, tag=None):
        """Agrega un mensaje al área de log con color verde o rojo."""
        self.log_text['state'] = "normal"
        self.log_text.insert(tk.END, message + "\n", tag)
        if tag == "success":
            self.log_text.tag_config("success", foreground="green")
        elif tag == "error":
            self.log_text.tag_config("error", foreground="red")
        self.log_text.see(tk.END)
        self.log_text['state'] = "disabled"
    
    def send_emails(self):
        """Envía correos a los destinatarios."""
        smtp_server = self.smtp_server_var.get().strip()
        smtp_port_str = self.smtp_port_var.get().strip()
        smtp_user = self.smtp_user_var.get().strip()
        smtp_password = self.smtp_password_var.get().strip()
        from_email = self.from_email_var.get().strip()
        subject = self.subject_var.get().strip()
        try:
            smtp_port = int(smtp_port_str)
        except ValueError:
            messagebox.showerror("Error", "El puerto SMTP debe ser un número.")
            return
        if not (smtp_server and smtp_port and smtp_user and smtp_password and from_email and subject):
            messagebox.showerror("Error", "Faltan datos de configuración SMTP o asunto.")
            return

        # Recoger datos de configuración SMTP y del mensaje
        smtp_server = self.smtp_server_var.get().strip()
        try:
            smtp_port = int(self.smtp_port_var.get().strip())
        except ValueError:
            messagebox.showerror("Error", "El puerto SMTP debe ser un número.")
            return
        smtp_user = self.smtp_user_var.get().strip()
        smtp_password = self.smtp_password_var.get().strip()
        from_email = self.from_email_var.get().strip()
        subject = self.subject_var.get().strip()
        message_body_template = self.message_text.get("1.0", tk.END).strip()
        
        # Validar que se hayan completado todos los campos
        if not smtp_server or not smtp_port or not smtp_user or not smtp_password or not from_email or not subject or not message_body_template:
            messagebox.showerror("Error", "Por favor, complete todos los campos de configuración y mensaje.")
            return
        
        # Obtener destinatarios:
        # Si ya se importó un CSV, se usa esa lista; de lo contrario, se parsea el texto ingresado (emails separados por comas)
        recipients_input = self.recipients_text.get("1.0", tk.END).strip()
        recipients_manual = []
        if recipients_input:
            for email in recipients_input.split(","):
                email = email.strip()
                if email:
                    # En ausencia de nombre, se usará el email
                    recipients_manual.append({"email": email, "nombre": email})
        
        if self.recipients:
            recipients_list = self.recipients
        else:
            recipients_list = recipients_manual
        
        if not recipients_list:
            messagebox.showerror("Error", "No se encontraron destinatarios.")
            return
        
        # Conectar al servidor SMTP
        try:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(smtp_user, smtp_password)
        except Exception as e:
            messagebox.showerror("Error", f"Error al conectar con el servidor SMTP: {e}")
            return
        
        success_count = 0
        # Enviar correo a cada destinatario
        for recipient in recipients_list:
            try:
                # Se reemplazan las variables {email} y {nombre} en el cuerpo del mensaje
                personalized_body = message_body_template.format(
                    email=recipient["email"],
                    nombre=recipient.get("nombre", recipient["email"])
                )
                msg = MIMEMultipart()
                msg["From"] = from_email
                msg["To"] = recipient["email"]
                msg["Subject"] = subject
                msg.attach(MIMEText(personalized_body, "plain"))
                
                server.send_message(msg)
                success_count += 1
                self.log(f"Correo enviado a: {recipient['email']}", "success")
            except Exception as e:
                self.log(f"Error al enviar a {recipient['email']}: {e}", "error")
        
        server.quit()
        self.log("Proceso completado. Correos enviados exitosamente: {}".format(success_count), "success")
        messagebox.showinfo("Información", "Proceso completado. Correos enviados: {}".format(success_count))

    def _on_canvas_configure(self, event):
        """Ajusta el ancho del frame scrollable cuando se redimensiona la ventana"""
        # Actualizar el ancho de la ventana del canvas para que coincida con el canvas
        self.canvas.itemconfig(self.canvas_window, width=event.width)

if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = EmailSenderGUI(root)
    root.mainloop()
