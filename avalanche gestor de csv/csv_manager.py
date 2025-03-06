import os
import csv
import sqlite3
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime
from PIL import Image, ImageTk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import ThemeDefinition

# Función para validar emails
def is_valid_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

# Función para determinar si un email es probablemente comercial basado en el dominio
def is_commercial_email(email):
    """Determina si un email es probablemente comercial basado en el dominio."""
    if not email or '@' not in email:
        return False
    
    domain = email.split('@')[1].lower()
    personal_domains = ['gmail.com', 'hotmail.com', 'yahoo.com', 'outlook.com', 'live.com', 
                        'icloud.com', 'aol.com', 'protonmail.com', 'mail.com', 'gmx.com']
    
    return domain not in personal_domains

# Clase de gestión de la base de datos
class DatabaseManager:
    def __init__(self, db_path="clients_database.db"):
        self.db_path = db_path
        self.conn = None
        self.create_database()
    
    def create_database(self):
        """Crea la base de datos si no existe y configura las tablas necesarias."""
        self.conn = sqlite3.connect(self.db_path)
        cursor = self.conn.cursor()
        
        # Tabla para clientes regulares
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY,
            name TEXT,
            email TEXT UNIQUE,
            imported_date TEXT,
            additional_info TEXT
        )
        ''')
        
        # Tabla para contactos comerciales
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS commercial_contacts (
            id INTEGER PRIMARY KEY,
            name TEXT,
            email TEXT UNIQUE,
            imported_date TEXT,
            company TEXT,
            additional_info TEXT
        )
        ''')
        
        # Tabla para emails no válidos
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS invalid_emails (
            id INTEGER PRIMARY KEY,
            email TEXT UNIQUE,
            name TEXT,
            imported_date TEXT,
            reason TEXT
        )
        ''')
        
        self.conn.commit()
    
    def close_connection(self):
        """Cierra la conexión a la base de datos."""
        if self.conn:
            self.conn.close()
    
    def add_client(self, name, email, additional_info=""):
        """Añade un cliente regular a la base de datos."""
        try:
            cursor = self.conn.cursor()
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                "INSERT OR IGNORE INTO clients (name, email, imported_date, additional_info) VALUES (?, ?, ?, ?)",
                (name, email, imported_date, additional_info)
            )
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al añadir cliente: {e}")
            return False
    
    def add_commercial_contact(self, name, email, company="", additional_info=""):
        """Añade un contacto comercial a la base de datos."""
        try:
            cursor = self.conn.cursor()
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                "INSERT OR IGNORE INTO commercial_contacts (name, email, imported_date, company, additional_info) VALUES (?, ?, ?, ?, ?)",
                (name, email, imported_date, company, additional_info)
            )
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al añadir contacto comercial: {e}")
            return False
    
    def add_invalid_email(self, email, name="", reason="Formato inválido"):
        """Añade un email no válido a la base de datos."""
        try:
            cursor = self.conn.cursor()
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                "INSERT OR IGNORE INTO invalid_emails (email, name, imported_date, reason) VALUES (?, ?, ?, ?)",
                (email, name, imported_date, reason)
            )
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al añadir email no válido: {e}")
            return False
    
    def mark_as_invalid(self, email, name="", reason="Marcado manualmente como inválido"):
        """Marca un email como no válido."""
        try:
            cursor = self.conn.cursor()
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                "INSERT OR IGNORE INTO invalid_emails (email, name, imported_date, reason) VALUES (?, ?, ?, ?)",
                (email, name, imported_date, reason)
            )
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al marcar email como inválido: {e}")
            return False
    
    def get_all_clients(self):
        """Obtiene todos los clientes regulares de la base de datos."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name, email, imported_date FROM clients")
        return cursor.fetchall()
    
    def get_all_commercial_contacts(self):
        """Obtiene todos los contactos comerciales de la base de datos."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT name, email, company, imported_date FROM commercial_contacts")
        return cursor.fetchall()
    
    def get_all_invalid_emails(self):
        """Obtiene todos los emails no válidos de la base de datos."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT email, name, imported_date, reason FROM invalid_emails")
        return cursor.fetchall()

    def delete_client(self, email):
        """Elimina un cliente de la base de datos."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM clients WHERE email=?", (email,))
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al eliminar cliente: {e}")
            return False
            
    def delete_commercial_contact(self, email):
        """Elimina un contacto comercial de la base de datos."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM commercial_contacts WHERE email=?", (email,))
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al eliminar contacto comercial: {e}")
            return False
            
    def delete_invalid_email(self, email):
        """Elimina un email no válido de la base de datos."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM invalid_emails WHERE email=?", (email,))
            self.conn.commit()
            return cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Error al eliminar email no válido: {e}")
            return False

# Clase para la interfaz gráfica
class CSVManagerApp:
    def __init__(self, root):
        self.root = root
        
        # Configuramos el tema base y personalizamos los colores
        self.setup_theme()
        
        self.root.title("Gestor de CSV - Clientes y Emails")
        self.root.geometry("950x650")
        self.root.resizable(True, True)
        
        # Configurar el icono de la aplicación
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error al cargar el icono: {e}")
        
        self.db_manager = DatabaseManager()
        
        self.setup_ui()
    
    def setup_theme(self):
        """Configura un tema con tonos marrones pastel."""
        # Usamos un tema base y personalizamos los colores manualmente
        # sandstone es un tema incorporado con tonos marrones
        self.style = ttk.Style(theme="sandstone")
        
        # Definimos colores personalizados
        primary_color = "#c8b6a6"      # Marrón claro pastel
        secondary_color = "#e8d5c4"    # Marrón más claro
        success_color = "#a39081"      # Marrón medio
        info_color = "#efe3d7"         # Beige claro
        warning_color = "#d7bda5"      # Marrón anaranjado
        danger_color = "#9b8579"       # Marrón oscuro
        light_color = "#f5efe9"        # Beige muy claro
        dark_color = "#7d6b5d"         # Marrón oscuro
        bg_color = "#f8f4ef"           # Fondo beige muy claro
        
        # Personalizamos algunos elementos específicos
        self.style.configure(".", font=("Helvetica", 10))
        
        # Mejoramos la visibilidad de los botones con colores más definidos y bordes
        self.style.configure("TButton", background=primary_color, foreground="black", borderwidth=1)
        self.style.configure("primary.TButton", background=primary_color, foreground="black")
        self.style.configure("secondary.TButton", background=secondary_color, foreground="black")
        self.style.configure("success.TButton", background=success_color, foreground="white")
        self.style.configure("info.TButton", background="#a2d2ff", foreground="black")  # Color más azulado para los botones de exportar
        self.style.configure("warning.TButton", background=warning_color, foreground="black")
        self.style.configure("danger.TButton", background="#e07a5f", foreground="white")  # Color rojo para el botón de eliminar
        
        # Mejoramos el contraste en estados de hover
        self.style.map("TButton", 
                     background=[('active', primary_color)],
                     foreground=[('active', 'black')])
        self.style.map("primary.TButton", 
                     background=[('active', primary_color)],
                     foreground=[('active', 'black')])
        self.style.map("success.TButton", 
                     background=[('active', success_color)],
                     foreground=[('active', 'white')])
        self.style.map("info.TButton", 
                     background=[('active', "#81b1e3")],  # Azul más oscuro en hover
                     foreground=[('active', 'black')])
        self.style.map("warning.TButton", 
                     background=[('active', warning_color)],
                     foreground=[('active', 'black')])
        self.style.map("danger.TButton", 
                     background=[('active', "#c15f48")],  # Rojo más oscuro en hover
                     foreground=[('active', 'white')])
        
        # Personalizamos algunos elementos del frame
        self.style.configure("TFrame", background=bg_color)
        self.style.configure("TLabel", background=bg_color, foreground=dark_color)
        self.style.configure("TLabelframe", background=bg_color, foreground=dark_color)
        self.style.configure("TLabelframe.Label", background=bg_color, foreground=dark_color)
    
    def setup_ui(self):
        """Configura la interfaz de usuario."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sección superior con logo y título
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        
        # Cargar y mostrar logo
        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
            logo_image = Image.open(logo_path)
            # Redimensionar logo si es necesario
            logo_image = logo_image.resize((100, 100), Image.Resampling.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = ttk.Label(header_frame, image=logo_photo)
            logo_label.image = logo_photo  # Mantener referencia para evitar que se elimine por el recolector de basura
            logo_label.pack(side=tk.LEFT, padx=10)
        except Exception as e:
            print(f"Error al cargar el logo: {e}")
        
        # Título de la aplicación
        title_label = ttk.Label(header_frame, text="Gestor de CSV - Clientes y Emails", 
                              font=("Helvetica", 18, "bold"))
        title_label.pack(side=tk.LEFT, padx=20, pady=10)
        
        # Sección de importación
        import_frame = ttk.Labelframe(main_frame, text="Importar CSV", padding="10")
        import_frame.pack(fill=tk.X, pady=10)
        
        import_buttons_frame = ttk.Frame(import_frame)
        import_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(import_buttons_frame, text="Seleccionar archivo CSV", 
                  command=self.import_csv, style="primary.TButton").pack(side=tk.LEFT, padx=5)
        
        self.auto_classify_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(import_buttons_frame, text="Clasificación automática", 
                       variable=self.auto_classify_var).pack(side=tk.LEFT, padx=15)
        
        # Notebook para las pestañas
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Pestaña de clientes
        self.clients_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.clients_frame, text="Clientes")
        
        # Pestaña de contactos comerciales
        self.commercial_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.commercial_frame, text="Contactos Comerciales")
        
        # Pestaña de emails no válidos
        self.invalid_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.invalid_frame, text="Emails No Válidos")
        
        # Configurar treeviews
        self.setup_clients_treeview()
        self.setup_commercial_treeview()
        self.setup_invalid_treeview()
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(action_frame, text="Actualizar Listas", 
                  command=self.update_lists, style="success.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Exportar Categoría Actual", 
                  command=self.export_current_category, style="info.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Exportar Todos los Datos", 
                  command=self.export_all_data, style="info.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Reclasificar Contacto", 
                  command=self.reclassify_contact, style="warning.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Eliminar Contacto(s)", 
                  command=self.delete_contacts, style="danger.TButton").pack(side=tk.LEFT, padx=5)
        
        # Footer con información de copyright
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=5)
        
        footer_text = "Diseñado con ❤️ por franHR\nCopyright © 2025 pcprogramacion.es"
        footer_label = ttk.Label(footer_frame, text=footer_text, justify=tk.CENTER, 
                                 font=("Helvetica", 9))
        footer_label.pack(pady=10)
        
        # Línea separadora para el footer
        ttk.Separator(main_frame).pack(fill=tk.X, before=footer_frame)
        
        # Cargar datos iniciales
        self.update_lists()
    
    def setup_clients_treeview(self):
        """Configura el treeview para clientes regulares."""
        columns = ("name", "email", "date")
        
        self.clients_tree = ttk.Treeview(self.clients_frame, columns=columns, show="headings", 
                                        selectmode="extended")
        self.clients_tree.heading("name", text="Nombre")
        self.clients_tree.heading("email", text="Email")
        self.clients_tree.heading("date", text="Fecha Importación")
        
        self.clients_tree.column("name", width=200)
        self.clients_tree.column("email", width=250)
        self.clients_tree.column("date", width=150)
        
        self.clients_tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.clients_frame, orient=tk.VERTICAL, 
                                 command=self.clients_tree.yview)
        self.clients_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_commercial_treeview(self):
        """Configura el treeview para contactos comerciales."""
        columns = ("name", "email", "company", "date")
        
        self.commercial_tree = ttk.Treeview(self.commercial_frame, columns=columns, show="headings", 
                                           selectmode="extended")
        self.commercial_tree.heading("name", text="Nombre")
        self.commercial_tree.heading("email", text="Email")
        self.commercial_tree.heading("company", text="Empresa")
        self.commercial_tree.heading("date", text="Fecha Importación")
        
        self.commercial_tree.column("name", width=180)
        self.commercial_tree.column("email", width=220)
        self.commercial_tree.column("company", width=180)
        self.commercial_tree.column("date", width=120)
        
        self.commercial_tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.commercial_frame, orient=tk.VERTICAL, 
                                 command=self.commercial_tree.yview)
        self.commercial_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_invalid_treeview(self):
        """Configura el treeview para emails no válidos."""
        columns = ("email", "name", "date", "reason")
        
        self.invalid_tree = ttk.Treeview(self.invalid_frame, columns=columns, show="headings", 
                                        selectmode="extended")
        self.invalid_tree.heading("email", text="Email")
        self.invalid_tree.heading("name", text="Nombre")
        self.invalid_tree.heading("date", text="Fecha Importación")
        self.invalid_tree.heading("reason", text="Motivo")
        
        self.invalid_tree.column("email", width=250)
        self.invalid_tree.column("name", width=200)
        self.invalid_tree.column("date", width=150)
        self.invalid_tree.column("reason", width=150)
        
        self.invalid_tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.invalid_frame, orient=tk.VERTICAL, 
                                 command=self.invalid_tree.yview)
        self.invalid_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def update_lists(self):
        """Actualiza los listados de clientes, comerciales y emails no válidos."""
        # Limpiar treeviews
        for item in self.clients_tree.get_children():
            self.clients_tree.delete(item)
        
        for item in self.commercial_tree.get_children():
            self.commercial_tree.delete(item)
        
        for item in self.invalid_tree.get_children():
            self.invalid_tree.delete(item)
        
        # Obtener y mostrar clientes
        clients = self.db_manager.get_all_clients()
        for client in clients:
            self.clients_tree.insert("", tk.END, values=client)
        
        # Obtener y mostrar contactos comerciales
        commercial_contacts = self.db_manager.get_all_commercial_contacts()
        for contact in commercial_contacts:
            self.commercial_tree.insert("", tk.END, values=contact)
        
        # Obtener y mostrar emails no válidos
        invalid_emails = self.db_manager.get_all_invalid_emails()
        for email in invalid_emails:
            self.invalid_tree.insert("", tk.END, values=email)
    
    def import_csv(self):
        """Importa datos desde un archivo CSV."""
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo CSV",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            new_clients = 0
            new_commercial = 0
            new_invalid = 0
            
            with open(filepath, 'r', encoding='utf-8-sig') as csv_file:
                csv_reader = csv.reader(csv_file)
                headers = next(csv_reader, None)  # Leer encabezados
                
                # Determinar índices de nombre, email y empresa
                name_index = 0
                email_index = 1
                company_index = -1
                
                if headers:
                    for i, header in enumerate(headers):
                        header_lower = header.lower()
                        if "nombre" in header_lower:
                            name_index = i
                        elif "email" in header_lower or "correo" in header_lower:
                            email_index = i
                        elif "empresa" in header_lower or "compañía" in header_lower or "company" in header_lower:
                            company_index = i
                
                # Procesar filas
                for row in csv_reader:
                    if len(row) > max(name_index, email_index):
                        name = row[name_index].strip()
                        email = row[email_index].strip().lower()
                        company = row[company_index].strip() if company_index >= 0 and company_index < len(row) else ""
                        
                        if not email:
                            continue
                            
                        # Verificar si el email ya existe en alguna tabla
                        existing_category = self.check_existing_email(email)
                        
                        if existing_category:
                            # El email ya existe, preguntar al usuario qué hacer
                            action = self.ask_duplicate_action(email, name, existing_category)
                            
                            if action == "skip":
                                # Omitir este email
                                continue
                            elif action == "client":
                                # Reclasificar a cliente
                                self.delete_from_all_categories(email)
                                if self.db_manager.add_client(name, email):
                                    new_clients += 1
                                continue
                            elif action == "commercial":
                                # Reclasificar a comercial
                                self.delete_from_all_categories(email)
                                if self.db_manager.add_commercial_contact(name, email, company):
                                    new_commercial += 1
                                continue
                            elif action == "invalid":
                                # Reclasificar a no válido
                                reason = "Marcado como no válido durante importación"
                                self.delete_from_all_categories(email)
                                if self.db_manager.add_invalid_email(email, name, reason):
                                    new_invalid += 1
                                continue
                        
                        # Procesamiento normal para emails que no existen
                        if is_valid_email(email):
                            # Es un email válido, determinamos si es cliente o comercial
                            auto_classify = self.auto_classify_var.get()
                            
                            if auto_classify and (is_commercial_email(email) or company):
                                # Email probablemente comercial o tiene empresa asociada
                                if self.db_manager.add_commercial_contact(name, email, company):
                                    new_commercial += 1
                            else:
                                # Cliente regular
                                if self.db_manager.add_client(name, email):
                                    new_clients += 1
                        elif email:
                            # Email con formato inválido
                            if self.db_manager.add_invalid_email(email, name):
                                new_invalid += 1
            
            self.update_lists()
            messagebox.showinfo(
                "Importación completada",
                f"Importación completada:\n"
                f"- {new_clients} nuevos clientes\n"
                f"- {new_commercial} nuevos contactos comerciales\n"
                f"- {new_invalid} nuevos emails no válidos"
            )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar el archivo: {str(e)}")
    
    def check_existing_email(self, email):
        """Verifica si un email ya existe en alguna categoría y devuelve en cuál."""
        cursor = self.db_manager.conn.cursor()
        
        # Verificar en clientes
        cursor.execute("SELECT COUNT(*) FROM clients WHERE email=?", (email,))
        if cursor.fetchone()[0] > 0:
            return "cliente"
        
        # Verificar en contactos comerciales
        cursor.execute("SELECT COUNT(*) FROM commercial_contacts WHERE email=?", (email,))
        if cursor.fetchone()[0] > 0:
            return "contacto comercial"
        
        # Verificar en emails no válidos
        cursor.execute("SELECT COUNT(*) FROM invalid_emails WHERE email=?", (email,))
        if cursor.fetchone()[0] > 0:
            return "email no válido"
        
        return None
    
    def ask_duplicate_action(self, email, name, existing_category):
        """Pregunta al usuario qué hacer con un email duplicado."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Email duplicado")
        dialog.geometry("750x250")  # Aumentado de 600x250 a 750x250 para más espacio horizontal
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg="#f8f4ef")
        
        # Variable para almacenar la respuesta
        result = tk.StringVar()
        
        # Crear contenido del diálogo
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Se ha encontrado un email duplicado:",
                 font=("Helvetica", 10, "bold")).pack(pady=(5, 10))
        ttk.Label(frame, text=f"{name} <{email}>").pack(pady=2)
        ttk.Label(frame, text=f"Este email ya existe como {existing_category}.").pack(pady=2)
        ttk.Label(frame, text="¿Qué desea hacer?").pack(pady=(10, 5))
        
        # Botones de acción
        actions_frame = ttk.Frame(frame)
        actions_frame.pack(pady=15)
        
        ttk.Button(actions_frame, text="Omitir", 
                  command=lambda: self.set_result_and_close(dialog, result, "skip"),
                  style="secondary.TButton").pack(side=tk.LEFT, padx=15)  # Aumentado el padx de 10 a 15
        
        ttk.Button(actions_frame, text="Importar como Cliente", 
                  command=lambda: self.set_result_and_close(dialog, result, "client"),
                  style="success.TButton").pack(side=tk.LEFT, padx=15)  # Aumentado el padx de 10 a 15
        
        ttk.Button(actions_frame, text="Importar como Comercial", 
                  command=lambda: self.set_result_and_close(dialog, result, "commercial"),
                  style="info.TButton").pack(side=tk.LEFT, padx=15)  # Aumentado el padx de 10 a 15
        
        ttk.Button(actions_frame, text="Marcar No Válido", 
                  command=lambda: self.set_result_and_close(dialog, result, "invalid"),
                  style="danger.TButton").pack(side=tk.LEFT, padx=15)  # Aumentado el padx de 10 a 15
        
        # Centrar la ventana de diálogo
        self.center_window(dialog)
        
        # Esperar a que el usuario responda
        self.root.wait_window(dialog)
        
        # Devolver el resultado
        return result.get() if result.get() else "skip"
    
    def set_result_and_close(self, dialog, result_var, value):
        """Establece el valor del resultado y cierra el diálogo."""
        result_var.set(value)
        dialog.destroy()
    
    def delete_from_all_categories(self, email):
        """Elimina un email de todas las categorías para evitar duplicados."""
        cursor = self.db_manager.conn.cursor()
        
        cursor.execute("DELETE FROM clients WHERE email=?", (email,))
        cursor.execute("DELETE FROM commercial_contacts WHERE email=?", (email,))
        cursor.execute("DELETE FROM invalid_emails WHERE email=?", (email,))
        
        self.db_manager.conn.commit()
    
    def reclassify_contact(self):
        """Reclasifica contactos entre las diferentes categorías."""
        # Determinar qué pestaña está activa
        current_tab = self.notebook.index(self.notebook.select())
        selected_items = []
        contact_type = ""
        contacts_to_reclassify = []
        
        # Obtener los contactos seleccionados de la pestaña activa
        if current_tab == 0:  # Pestaña de clientes
            selected_items = self.clients_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.clients_tree.item(item, 'values')
                    contacts_to_reclassify.append({"name": values[0], "email": values[1]})
                contact_type = "cliente"
        elif current_tab == 1:  # Pestaña de contactos comerciales
            selected_items = self.commercial_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.commercial_tree.item(item, 'values')
                    # Para contactos comerciales, también guardamos la empresa
                    contacts_to_reclassify.append({"name": values[0], "email": values[1], "company": values[2]})
                contact_type = "contacto comercial"
        elif current_tab == 2:  # Pestaña de emails no válidos
            selected_items = self.invalid_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.invalid_tree.item(item, 'values')
                    contacts_to_reclassify.append({"email": values[0], "name": values[1]})
                contact_type = "email no válido"
        
        if not selected_items:
            messagebox.showinfo("Información", "Por favor, seleccione al menos un contacto para reclasificar.")
            return
        
        # Crear ventana de diálogo para seleccionar la nueva categoría
        target_category_window = tk.Toplevel(self.root)
        target_category_window.title("Reclasificar Contactos")
        target_category_window.geometry("400x220")  # Aumentado de 180 a 220 para más espacio vertical
        target_category_window.resizable(False, False)
        target_category_window.transient(self.root)
        target_category_window.grab_set()
        target_category_window.configure(bg="#f8f4ef")  # Usar el color de fondo beige
        
        # Aplicar el mismo estilo a la ventana emergente
        frame = ttk.Frame(target_category_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        num_selected = len(selected_items)
        plural = "s" if num_selected > 1 else ""
        ttk.Label(frame, 
                text=f"Reclasificar {num_selected} {contact_type}{plural} seleccionado{plural} a:", 
                justify=tk.CENTER).pack(pady=10)
        
        # Mostrar los primeros contactos (hasta 3) como referencia
        contact_display = "\n".join([f"{c.get('name', '')} <{c.get('email', '')}>".strip() for c in contacts_to_reclassify[:3]])
        if num_selected > 3:
            contact_display += f"\n... y {num_selected - 3} más"
            
        ttk.Label(frame, text=contact_display, justify=tk.LEFT).pack(pady=5)
        
        # Variable para almacenar la selección
        target_category = tk.StringVar()
        
        # Crear opciones de categoría, excluyendo la categoría actual
        categories = []
        if contact_type != "cliente":
            categories.append(("Cliente", 0))
        if contact_type != "contacto comercial":
            categories.append(("Contacto Comercial", 1))
        if contact_type != "email no válido":
            categories.append(("Email No Válido", 2))
        
        # Frame para los botones de categoría
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(pady=10)
        
        # Función para manejar la selección de categoría
        def on_category_select(category_index):
            target_category.set(str(category_index))
            target_category_window.destroy()
        
        # Crear botones para cada categoría disponible
        for cat_name, cat_index in categories:
            ttk.Button(buttons_frame, text=cat_name, 
                     command=lambda idx=cat_index: on_category_select(idx)).pack(side=tk.LEFT, padx=5)
        
        # Botón para cancelar
        ttk.Button(frame, text="Cancelar", 
                 command=target_category_window.destroy).pack(pady=5)
        
        # Centrar la ventana en la pantalla
        self.center_window(target_category_window)
        
        # Esperar a que se cierre la ventana
        self.root.wait_window(target_category_window)
        
        # Si no se seleccionó ninguna categoría, retornar
        if not target_category.get():
            return
        
        # Procesar la reclasificación según la categoría seleccionada para todos los contactos
        new_category_index = int(target_category.get())
        
        # Si la nueva categoría es "Email No Válido", preguntar el motivo una sola vez
        reason = None
        if new_category_index == 2 and contact_type != "email no válido":
            reason = simpledialog.askstring("Motivo", 
                                          "Ingrese el motivo por el que estos emails son considerados no válidos:",
                                          parent=self.root) or "Marcado manualmente como inválido"
        
        # Procesar cada contacto
        success_count = 0
        for contact in contacts_to_reclassify:
            if self._process_reclassification(contact_type, contact, new_category_index, reason):
                success_count += 1
        
        # Actualizar las listas después de procesar todos los contactos
        self.update_lists()
        
        # Mostrar mensaje de éxito
        target_types = {0: "clientes", 1: "contactos comerciales", 2: "emails no válidos"}
        messagebox.showinfo("Éxito", f"{success_count} contacto(s) reclasificado(s) correctamente como {target_types[new_category_index]}.")
    
    def _process_reclassification(self, current_type, contact, new_category_index, provided_reason=None):
        """Procesa la reclasificación de un contacto a una nueva categoría."""
        try:
            name = contact.get("name", "")
            email = contact.get("email", "")
            company = contact.get("company", "")
            
            cursor = self.db_manager.conn.cursor()
            
            # Eliminar de la categoría actual
            if current_type == "cliente":
                cursor.execute("DELETE FROM clients WHERE email=?", (email,))
            elif current_type == "contacto comercial":
                cursor.execute("DELETE FROM commercial_contacts WHERE email=?", (email,))
            elif current_type == "email no válido":
                cursor.execute("DELETE FROM invalid_emails WHERE email=?", (email,))
            
            # Insertar en la nueva categoría
            if new_category_index == 0:  # Cliente
                self.db_manager.add_client(name, email)
            elif new_category_index == 1:  # Contacto comercial
                self.db_manager.add_commercial_contact(name, email, company)
            elif new_category_index == 2:  # Email no válido
                reason = provided_reason or "Marcado manualmente como inválido"
                self.db_manager.add_invalid_email(email, name, reason)
            
            self.db_manager.conn.commit()
            return True
            
        except Exception as e:
            print(f"Error al reclasificar el contacto {email}: {str(e)}")
            return False
    
    def center_window(self, window):
        """Centra una ventana en la pantalla."""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')
    
    def export_current_category(self):
        """Exporta solo los datos de la categoría actualmente seleccionada."""
        current_tab = self.notebook.index(self.notebook.select())
        
        if current_tab == 0:  # Clientes
            self._export_category("clientes", self.db_manager.get_all_clients(),
                                 ["Nombre", "Email", "Fecha Importación"])
        elif current_tab == 1:  # Comerciales
            self._export_category("contactos_comerciales", self.db_manager.get_all_commercial_contacts(),
                                 ["Nombre", "Email", "Empresa", "Fecha Importación"])
        elif current_tab == 2:  # No válidos
            self._export_category("emails_no_validos", self.db_manager.get_all_invalid_emails(),
                                 ["Email", "Nombre", "Fecha Importación", "Motivo"])
    
    def _export_category(self, filename_prefix, data, headers):
        """Exporta una categoría específica a un archivo CSV."""
        filepath = filedialog.asksaveasfilename(
            title=f"Guardar {filename_prefix}",
            defaultextension=".csv",
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
            initialfile=f"{filename_prefix}.csv"
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(data)
            
            messagebox.showinfo(
                "Exportación completada",
                f"Los datos han sido exportados exitosamente a:\n{filepath}"
            )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar los datos: {str(e)}")
    
    def export_all_data(self):
        """Exporta todos los datos a archivos CSV."""
        export_dir = filedialog.askdirectory(title="Seleccionar carpeta para exportación de todos los datos")
        
        if not export_dir:
            return
        
        try:
            # Exportar clientes
            clients_path = os.path.join(export_dir, "clientes.csv")
            with open(clients_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Nombre", "Email", "Fecha Importación"])
                
                clients = self.db_manager.get_all_clients()
                writer.writerows(clients)
            
            # Exportar contactos comerciales
            commercial_path = os.path.join(export_dir, "contactos_comerciales.csv")
            with open(commercial_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)  # Corregido: era writer = csv.writer.f)
                writer.writerow(["Nombre", "Email", "Empresa", "Fecha Importación"])
                
                commercial_contacts = self.db_manager.get_all_commercial_contacts()
                writer.writerows(commercial_contacts)
            
            # Exportar emails no válidos
            invalid_path = os.path.join(export_dir, "emails_no_validos.csv")
            with open(invalid_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["Email", "Nombre", "Fecha Importación", "Motivo"])
                
                invalid_emails = self.db_manager.get_all_invalid_emails()
                writer.writerows(invalid_emails)
            
            messagebox.showinfo(
                "Exportación completada",
                f"Los datos han sido exportados exitosamente a:\n"
                f"- {clients_path}\n"
                f"- {commercial_path}\n"
                f"- {invalid_path}"
            )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar los datos: {str(e)}")
    
    def delete_contacts(self):
        """Elimina los contactos seleccionados de la categoría actual."""
        # Determinar qué pestaña está activa
        current_tab = self.notebook.index(self.notebook.select())
        selected_items = []
        contact_type = ""
        contacts_to_delete = []
        
        # Obtener los contactos seleccionados de la pestaña activa
        if current_tab == 0:  # Pestaña de clientes
            selected_items = self.clients_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.clients_tree.item(item, 'values')
                    contacts_to_delete.append({"name": values[0], "email": values[1]})
                contact_type = "cliente"
        elif current_tab == 1:  # Pestaña de contactos comerciales
            selected_items = self.commercial_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.commercial_tree.item(item, 'values')
                    contacts_to_delete.append({"name": values[0], "email": values[1]})
                contact_type = "contacto comercial"
        elif current_tab == 2:  # Pestaña de emails no válidos
            selected_items = self.invalid_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.invalid_tree.item(item, 'values')
                    contacts_to_delete.append({"email": values[0], "name": values[1]})
                contact_type = "email no válido"
        
        if not selected_items:
            messagebox.showinfo("Información", "Por favor, seleccione al menos un contacto para eliminar.")
            return
        
        # Pedir confirmación antes de eliminar
        num_selected = len(selected_items)
        plural = "s" if num_selected > 1 else ""
        if not messagebox.askyesno("Confirmar eliminación", 
                                  f"¿Está seguro que desea eliminar {num_selected} {contact_type}{plural} seleccionado{plural}?"):
            return
        
        # Procesar la eliminación de cada contacto
        deleted_count = 0
        for contact in contacts_to_delete:
            email = contact.get("email", "")
            
            if current_tab == 0:  # Cliente
                if self.db_manager.delete_client(email):
                    deleted_count += 1
            elif current_tab == 1:  # Contacto comercial
                if self.db_manager.delete_commercial_contact(email):
                    deleted_count += 1
            elif current_tab == 2:  # Email no válido
                if self.db_manager.delete_invalid_email(email):
                    deleted_count += 1
        
        # Actualizar la vista después de eliminar
        self.update_lists()
        
        # Mostrar mensaje de éxito
        messagebox.showinfo("Eliminación completada", 
                          f"Se han eliminado {deleted_count} contacto(s) correctamente.")
    
    def on_closing(self):
        """Maneja el cierre de la aplicación."""
        self.db_manager.close_connection()
        self.root.destroy()

# Ejecutar la aplicación
if __name__ == "__main__":
    root = ttk.Window(themename="sandstone")
    app = CSVManagerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
