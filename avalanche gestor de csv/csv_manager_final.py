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
import sys
import traceback

# Función para obtener la ruta base de la aplicación
def get_application_path():
    """Obtiene la ruta base de la aplicación, ya sea en modo desarrollo o ejecutable"""
    if getattr(sys, 'frozen', False):
        # Si estamos en un ejecutable creado con PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # En modo desarrollo
        return os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    """Obtiene la ruta absoluta a un recurso, funciona tanto para desarrollo como para PyInstaller"""
    try:
        # PyInstaller crea un directorio temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = get_application_path()
    
    return os.path.join(base_path, relative_path)

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
        # Aseguramos que la base de datos se cree en la misma carpeta que la aplicación
        self.db_path = os.path.join(get_application_path(), db_path)
        self.conn = None
        self.create_database()
        # Bandera para habilitar/deshabilitar el modo de depuración
        self.debug_mode = False
    
    def debug_comparison(self, field_name, original, nuevo, normalizado_original, normalizado_nuevo):
        """Muestra información de depuración sobre la comparación de valores."""
        if self.debug_mode:
            print(f"DEBUG - Campo: {field_name}")
            print(f"  Original DB: '{original}' (tipo: {type(original).__name__})")
            print(f"  Nuevo: '{nuevo}' (tipo: {type(nuevo).__name__})")
            print(f"  Original Normalizado: '{normalizado_original}'")
            print(f"  Nuevo Normalizado: '{normalizado_nuevo}'")
            print(f"  ¿Son diferentes?: {normalizado_original != normalizado_nuevo}")
            print(f"  Longitudes: Original={len(str(normalizado_original))}, Nuevo={len(str(normalizado_nuevo))}")
            print("-------------------------------------------")
    
    def set_debug_mode(self, enabled=True):
        """Habilita o deshabilita el modo de depuración."""
        self.debug_mode = enabled
        return self.debug_mode
    
    def create_database(self):
        """Crea la base de datos si no existe y configura las tablas necesarias."""
        # Verificar si el directorio de la base de datos existe
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            try:
                os.makedirs(db_dir)
            except Exception as e:
                print(f"Error al crear el directorio para la base de datos: {e}")
        
        # Conectar a la base de datos (la crea si no existe)
        try:
            self.conn = sqlite3.connect(self.db_path)
            cursor = self.conn.cursor()
            
            # Tabla para clientes regulares
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY,
                name TEXT,
                email TEXT UNIQUE,
                imported_date TEXT,
                client_code TEXT,
                address TEXT,
                postal_code TEXT,
                town TEXT,
                city TEXT,
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
                client_code TEXT,
                address TEXT,
                postal_code TEXT,
                town TEXT,
                city TEXT,
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
            
            # Tabla para guardar configuraciones
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS config (
                id INTEGER PRIMARY KEY,
                key TEXT UNIQUE,
                value TEXT
            )
            ''')
            
            # Verificar si las nuevas columnas existen en la tabla clients y añadirlas si faltan
            # Primero obtener la información actual de la tabla clients
            cursor.execute("PRAGMA table_info(clients)")
            existing_columns = [column[1] for column in cursor.fetchall()]
            
            # Añadir columnas nuevas si no existen
            if 'client_code' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN client_code TEXT")
            if 'address' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN address TEXT")
            if 'postal_code' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN postal_code TEXT")
            if 'town' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN town TEXT")
            if 'city' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN city TEXT")
            if 'additional_info' not in existing_columns:
                cursor.execute("ALTER TABLE clients ADD COLUMN additional_info TEXT")
            
            # Verificar y añadir columnas a la tabla commercial_contacts
            cursor.execute("PRAGMA table_info(commercial_contacts)")
            existing_columns = [column[1] for column in cursor.fetchall()]
            
            if 'client_code' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN client_code TEXT")
            if 'address' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN address TEXT")
            if 'postal_code' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN postal_code TEXT")
            if 'town' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN town TEXT")
            if 'city' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN city TEXT")
            if 'additional_info' not in existing_columns:
                cursor.execute("ALTER TABLE commercial_contacts ADD COLUMN additional_info TEXT")
            
            self.conn.commit()
            print(f"Base de datos creada/actualizada exitosamente en {self.db_path}")
            
        except sqlite3.Error as e:
            print(f"Error al crear/conectar a la base de datos: {e}")
    
    def close_connection(self):
        """Cierra la conexión a la base de datos."""
        if self.conn:
            self.conn.close()
    
    def add_client(self, name, email, client_code="", address="", postal_code="", town="", city="", additional_info="", imported_date=None):
        """Añade un nuevo cliente a la base de datos."""
        if not self.conn:
            self.create_database()
        
        # Asegurar que tenemos una fecha para imported_date
        if imported_date is None:
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Normalizar valores: convertir None a cadenas vacías y eliminar espacios en blanco al inicio y final
        name = (name or "").strip()
        email = (email or "").strip().lower()
        client_code = (client_code or "").strip()
        address = (address or "").strip()
        postal_code = (postal_code or "").strip()
        town = (town or "").strip()
        city = (city or "").strip()
        additional_info = (additional_info or "").strip()
        
        # Comprobar si el email ya existe
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, name, client_code, address, postal_code, town, city, additional_info FROM clients WHERE email = ?", (email,))
        existing_client = cursor.fetchone()
        
        try:
            if existing_client:
                # Cliente ya existe, verificar si hay cambios
                client_id = existing_client[0]
                # Normalizar los valores existentes de la misma manera que los nuevos
                existing_name = (existing_client[1] or "").strip()
                existing_client_code = (existing_client[2] or "").strip()
                existing_address = (existing_client[3] or "").strip()
                existing_postal_code = (existing_client[4] or "").strip()
                existing_town = (existing_client[5] or "").strip()
                existing_city = (existing_client[6] or "").strip()
                existing_additional_info = (existing_client[7] or "").strip() if len(existing_client) > 7 else ""
                
                # Función para normalizar aún más los valores para comparación (elimina espacios múltiples, etc.)
                def normalize_deeper(value):
                    if value is None:
                        return ""
                    # Convertir a string si no lo es
                    value = str(value).strip()
                    # Eliminar caracteres especiales que puedan interferir en la comparación
                    value = value.replace('*', '').replace('\t', ' ').replace('\r', '').replace('\n', ' ')
                    # Eliminar espacios múltiples y normalizar
                    return ' '.join(value.split())
                
                # Aplicar normalización más profunda a todos los valores
                name_norm = normalize_deeper(name)
                existing_name_norm = normalize_deeper(existing_name)
                
                client_code_norm = normalize_deeper(client_code)
                existing_client_code_norm = normalize_deeper(existing_client_code)
                
                address_norm = normalize_deeper(address)
                existing_address_norm = normalize_deeper(existing_address)
                
                postal_code_norm = normalize_deeper(postal_code)
                existing_postal_code_norm = normalize_deeper(existing_postal_code)
                
                town_norm = normalize_deeper(town)
                existing_town_norm = normalize_deeper(existing_town)
                
                city_norm = normalize_deeper(city)
                existing_city_norm = normalize_deeper(existing_city)
                
                # Verificar si algún campo ha cambiado (usando valores normalizados para la comparación)
                changes = []
                if name_norm != existing_name_norm:
                    self.debug_comparison("nombre", existing_name, name, existing_name_norm, name_norm)
                    changes.append(f"nombre: '{existing_name}' -> '{name}'")
                
                if client_code_norm != existing_client_code_norm:
                    self.debug_comparison("código cliente", existing_client_code, client_code, existing_client_code_norm, client_code_norm)
                    changes.append(f"código cliente: '{existing_client_code}' -> '{client_code}'")
                
                if address_norm != existing_address_norm:
                    self.debug_comparison("dirección", existing_address, address, existing_address_norm, address_norm)
                    changes.append(f"dirección: '{existing_address}' -> '{address}'")
                
                if postal_code_norm != existing_postal_code_norm:
                    self.debug_comparison("código postal", existing_postal_code, postal_code, existing_postal_code_norm, postal_code_norm)
                    changes.append(f"código postal: '{existing_postal_code}' -> '{postal_code}'")
                
                if town_norm != existing_town_norm:
                    self.debug_comparison("población", existing_town, town, existing_town_norm, town_norm)
                    changes.append(f"población: '{existing_town}' -> '{town}'")
                
                if city_norm != existing_city_norm:
                    self.debug_comparison("ciudad", existing_city, city, existing_city_norm, city_norm)
                    changes.append(f"ciudad: '{existing_city}' -> '{city}'")
                
                additional_info_norm = normalize_deeper(additional_info)
                existing_additional_info_norm = normalize_deeper(existing_additional_info)
                if additional_info_norm != existing_additional_info_norm:
                    self.debug_comparison("información adicional", existing_additional_info, additional_info, existing_additional_info_norm, additional_info_norm)
                    changes.append(f"información adicional: '{existing_additional_info}' -> '{additional_info}'")
                
                if changes:
                    # Hay cambios, actualizar el cliente
                    cursor.execute(
                        "UPDATE clients SET name = ?, client_code = ?, address = ?, postal_code = ?, town = ?, city = ?, additional_info = ? WHERE id = ?",
                        (name, client_code, address, postal_code, town, city, additional_info, client_id)
                    )
                    self.conn.commit()
                    return {"updated": True, "changes": changes}  # Retornar que fue actualizado con los cambios
                else:
                    # No hay cambios, no hacer nada
                    return {"updated": False, "changes": []}  # Retornar que no hubo cambios
            else:
                # Nuevo cliente, insertarlo
                cursor.execute(
                    "INSERT INTO clients (name, email, imported_date, client_code, address, postal_code, town, city, additional_info) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (name, email, imported_date, client_code, address, postal_code, town, city, additional_info)
                )
                self.conn.commit()
                return {"new": True}  # Retornar que es nuevo
                
        except sqlite3.IntegrityError as e:
            print(f"Error de integridad al añadir cliente: {e}")
            return {"error": str(e)}
        except Exception as e:
            print(f"Error al añadir cliente: {e}")
            return {"error": str(e)}
    
    def add_commercial_contact(self, name, email, company="", client_code="", address="", postal_code="", town="", city="", additional_info=""):
        """Añade un contacto comercial a la base de datos."""
        try:
            if not self.conn:
                self.create_database()
                
            # Normalizar valores para evitar problemas
            name = str(name).strip() if name else ""
            email = str(email).strip().lower() if email else ""
            company = str(company).strip() if company else ""
            client_code = str(client_code).strip() if client_code else ""
            address = str(address).strip() if address else ""
            postal_code = str(postal_code).strip() if postal_code else ""
            town = str(town).strip() if town else ""
            city = str(city).strip() if city else ""
            additional_info = str(additional_info).strip() if additional_info else ""
            
            cursor = self.conn.cursor()
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Verificar si necesitamos actualizar en lugar de insertar
            cursor.execute("SELECT COUNT(*) FROM commercial_contacts WHERE email = ?", (email,))
            count = cursor.fetchone()[0]
            
            if count > 0:
                # Actualizar contacto existente
                cursor.execute(
                    """UPDATE commercial_contacts 
                       SET name = ?, company = ?, client_code = ?, address = ?, 
                           postal_code = ?, town = ?, city = ?, additional_info = ? 
                       WHERE email = ?""",
                    (name, company, client_code, address, postal_code, town, city, additional_info, email)
                )
                self.conn.commit()
                return {"updated": True}
            else:
                # Insertar nuevo contacto
                cursor.execute(
                    """INSERT INTO commercial_contacts 
                       (name, email, imported_date, company, client_code, address, postal_code, town, city, additional_info) 
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (name, email, imported_date, company, client_code, address, postal_code, town, city, additional_info)
                )
                self.conn.commit()
                return {"new": True}
                
        except sqlite3.IntegrityError as e:
            print(f"Error de integridad al añadir contacto comercial: {e}")
            return {"error": str(e)}
        except Exception as e:
            print(f"Error al añadir contacto comercial: {e}")
            return {"error": str(e)}
    
    def add_commercial_contact_with_changes(self, name, email, company="", client_code="", address="", postal_code="", town="", city="", additional_info="", imported_date=None):
        """Añade o actualiza un contacto comercial y detecta cambios."""
        if not self.conn:
            self.create_database()
        
        # Asegurar que tenemos una fecha para imported_date
        if imported_date is None:
            imported_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Normalizar valores: convertir None a cadenas vacías y eliminar espacios
        name = (name or "").strip()
        email = (email or "").strip().lower()
        company = (company or "").strip()
        client_code = (client_code or "").strip()
        address = (address or "").strip()
        postal_code = (postal_code or "").strip()
        town = (town or "").strip()
        city = (city or "").strip()
        additional_info = (additional_info or "").strip()
        
        # Comprobar si el email ya existe en la tabla de contactos comerciales
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, name, company, client_code, address, postal_code, town, city, additional_info FROM commercial_contacts WHERE email = ?", (email,))
        existing_contact = cursor.fetchone()
        
        try:
            if existing_contact:
                # Contacto ya existe, verificar si hay cambios
                contact_id = existing_contact[0]
                # Normalizar los valores existentes
                existing_name = (existing_contact[1] or "").strip()
                existing_company = (existing_contact[2] or "").strip()
                existing_client_code = (existing_contact[3] or "").strip()
                existing_address = (existing_contact[4] or "").strip()
                existing_postal_code = (existing_contact[5] or "").strip()
                existing_town = (existing_contact[6] or "").strip()
                existing_city = (existing_contact[7] or "").strip()
                existing_additional_info = (existing_contact[8] or "").strip()
                
                # Función para normalizar aún más los valores para comparación
                def normalize_deeper(value):
                    if value is None:
                        return ""
                    value = str(value).strip()
                    value = value.replace('*', '').replace('\t', ' ').replace('\r', '').replace('\n', ' ')
                    return ' '.join(value.split())
                
                # Aplicar normalización más profunda a todos los valores
                name_norm = normalize_deeper(name)
                existing_name_norm = normalize_deeper(existing_name)
                company_norm = normalize_deeper(company)
                existing_company_norm = normalize_deeper(existing_company)
                client_code_norm = normalize_deeper(client_code)
                existing_client_code_norm = normalize_deeper(existing_client_code)
                address_norm = normalize_deeper(address)
                existing_address_norm = normalize_deeper(existing_address)
                postal_code_norm = normalize_deeper(postal_code)
                existing_postal_code_norm = normalize_deeper(existing_postal_code)
                town_norm = normalize_deeper(town)
                existing_town_norm = normalize_deeper(existing_town)
                city_norm = normalize_deeper(city)
                existing_city_norm = normalize_deeper(existing_city)
                additional_info_norm = normalize_deeper(additional_info)
                existing_additional_info_norm = normalize_deeper(existing_additional_info)
                
                # Verificar si algún campo ha cambiado (usando valores normalizados)
                changes = []
                
                # Verificar cambios en el nombre (ignorando diferencias solo de mayúsculas/minúsculas)
                if name_norm.lower() != existing_name_norm.lower():
                    self.debug_comparison("nombre", existing_name, name, existing_name_norm, name_norm)
                    changes.append(f"nombre: '{existing_name}' -> '{name}'")
                
                # Verificar cambios en otros campos
                if company_norm != existing_company_norm:
                    self.debug_comparison("empresa", existing_company, company, existing_company_norm, company_norm)
                    changes.append(f"empresa: '{existing_company}' -> '{company}'")
                
                if client_code_norm != existing_client_code_norm:
                    self.debug_comparison("código cliente", existing_client_code, client_code, existing_client_code_norm, client_code_norm)
                    changes.append(f"código cliente: '{existing_client_code}' -> '{client_code}'")
                
                if address_norm != existing_address_norm:
                    self.debug_comparison("dirección", existing_address, address, existing_address_norm, address_norm)
                    changes.append(f"dirección: '{existing_address}' -> '{address}'")
                
                if postal_code_norm != existing_postal_code_norm:
                    self.debug_comparison("código postal", existing_postal_code, postal_code, existing_postal_code_norm, postal_code_norm)
                    changes.append(f"código postal: '{existing_postal_code}' -> '{postal_code}'")
                
                if town_norm != existing_town_norm:
                    self.debug_comparison("población", existing_town, town, existing_town_norm, town_norm)
                    changes.append(f"población: '{existing_town}' -> '{town}'")
                
                if city_norm != existing_city_norm:
                    self.debug_comparison("ciudad", existing_city, city, existing_city_norm, city_norm)
                    changes.append(f"ciudad: '{existing_city}' -> '{city}'")
                
                additional_info_norm = normalize_deeper(additional_info)
                existing_additional_info_norm = normalize_deeper(existing_additional_info)
                if additional_info_norm != existing_additional_info_norm:
                    self.debug_comparison("información adicional", existing_additional_info, additional_info, existing_additional_info_norm, additional_info_norm)
                    changes.append(f"información adicional: '{existing_additional_info}' -> '{additional_info}'")
                
                if changes:
                    # Solo actualizar si hay cambios reales
                    cursor.execute(
                        """UPDATE commercial_contacts 
                           SET name = ?, company = ?, client_code = ?, address = ?, 
                               postal_code = ?, town = ?, city = ?, additional_info = ? 
                           WHERE id = ?""",
                        (name, company, client_code, address, postal_code, town, city, additional_info, contact_id)
                    )
                    self.conn.commit()
                    return {"updated": True, "changes": changes}
                else:
                    return {"updated": False, "changes": []}
            else:
                # Nuevo contacto, insertarlo
                cursor.execute(
                    """INSERT INTO commercial_contacts 
                       (name, email, imported_date, company, client_code, address, postal_code, town, city, additional_info) 
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (name, email, imported_date, company, client_code, address, postal_code, town, city, additional_info)
                )
                self.conn.commit()
                return {"new": True}
                
        except sqlite3.IntegrityError as e:
            print(f"Error de integridad al añadir contacto comercial: {e}")
            return {"error": str(e)}
        except Exception as e:
            print(f"Error al añadir contacto comercial: {e}")
            return {"error": str(e)}
    
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
            return True
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
            return True
        except sqlite3.Error as e:
            print(f"Error al marcar email como inválido: {e}")
            return False
    
    def get_all_clients(self):
        """Obtiene todos los clientes regulares de la base de datos."""
        cursor = self.conn.cursor()
        # Verificar qué columnas existen en la tabla
        cursor.execute("PRAGMA table_info(clients)")
        columns = [column[1] for column in cursor.fetchall()]
        
        # Construir la consulta SQL en función de las columnas existentes
        base_columns = ["name", "email", "imported_date"]
        extra_columns = ["client_code", "address", "postal_code", "town", "city", "additional_info"]
        
        # Verificar qué columnas extra existen
        select_columns = base_columns + [col for col in extra_columns if col in columns]
        
        # Construir y ejecutar la consulta
        query = f"SELECT {', '.join(select_columns)} FROM clients"
        cursor.execute(query)
        
        # Obtener los resultados
        results = cursor.fetchall()
        
        # Si faltan columnas en los resultados, añadir valores vacíos
        if len(select_columns) < len(base_columns) + len(extra_columns):
            full_results = []
            for row in results:
                # Convertir la tupla a lista para poder modificarla
                row_list = list(row)
                
                # Añadir valores vacíos para las columnas que faltan
                for col in base_columns + extra_columns:
                    if col not in select_columns:
                        # Determinar la posición donde debería ir esta columna
                        all_columns = base_columns + extra_columns
                        position = all_columns.index(col)
                        # Asegurarnos de no exceder el tamaño de la lista
                        while len(row_list) <= position:
                            row_list.append("")
                        # Insertar el valor vacío en la posición correcta
                        if position < len(row_list):
                            row_list.insert(position, "")
                
                # Convertir de nuevo a tupla
                full_results.append(tuple(row_list))
            return full_results
        
        return results
    
    def get_all_commercial_contacts(self):
        """Obtiene todos los contactos comerciales de la base de datos."""
        cursor = self.conn.cursor()
        # Verificar qué columnas existen en la tabla
        cursor.execute("PRAGMA table_info(commercial_contacts)")
        columns = [column[1] for column in cursor.fetchall()]
        
        # Construir la consulta SQL en función de las columnas existentes
        base_columns = ["name", "email", "company", "imported_date"]
        extra_columns = ["client_code", "address", "postal_code", "town", "city", "additional_info"]
        
        # Verificar qué columnas extra existen
        select_columns = base_columns + [col for col in extra_columns if col in columns]
        
        # Construir y ejecutar la consulta
        query = f"SELECT {', '.join(select_columns)} FROM commercial_contacts"
        cursor.execute(query)
        
        # Obtener los resultados
        results = cursor.fetchall()
        
        # Si faltan columnas en los resultados, añadir valores vacíos
        if len(select_columns) < len(base_columns) + len(extra_columns):
            full_results = []
            for row in results:
                # Convertir la tupla a lista para poder modificarla
                row_list = list(row)
                
                # Añadir valores vacíos para las columnas que faltan
                for col in base_columns + extra_columns:
                    if col not in select_columns:
                        # Determinar la posición donde debería ir esta columna
                        all_columns = base_columns + extra_columns
                        position = all_columns.index(col)
                        # Asegurarnos de no exceder el tamaño de la lista
                        while len(row_list) <= position:
                            row_list.append("")
                        # Insertar el valor vacío en la posición correcta
                        if position < len(row_list):
                            row_list.insert(position, "")
                
                # Convertir de nuevo a tupla
                full_results.append(tuple(row_list))
            return full_results
        
        return results
    
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

    def save_config(self, key, value):
        """Guarda una configuración en la base de datos."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("INSERT OR REPLACE INTO config (key, value) VALUES (?, ?)", (key, value))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al guardar configuración: {e}")
            return False
    
    def get_config(self, key, default=None):
        """Obtiene una configuración de la base de datos."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT value FROM config WHERE key = ?", (key,))
            result = cursor.fetchone()
            return result[0] if result else default
        except sqlite3.Error as e:
            print(f"Error al obtener configuración: {e}")
            return default

# Clase para la interfaz gráfica
class CSVManagerApp:
    def __init__(self, root):
        self.root = root
        
        # Configuramos el tema base y personalizamos los colores
        self.setup_theme()
        
        self.root.title("Gestor de CSV - Clientes y Emails")
        # Aumentar el tamaño de la ventana principal (era 1400x800)
        self.root.geometry("1400x800")
        self.root.resizable(True, True)
        
        # Configurar el icono de la aplicación
        try:
            icon_path = resource_path("logo.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Error al cargar el icono: {e}")
        
        # Asegurarnos de que el archivo de base de datos existe antes de continuar
        db_path = os.path.join(get_application_path(), "clients_database.db")
        self.db_manager = DatabaseManager(db_path)
        
        # Lista para almacenar los IDs de contactos nuevos
        self.new_contacts = []
        # Variable para controlar si se omiten todos los duplicados
        self.skip_all_duplicates = False
        # Lista para almacenar los contactos modificados y sus campos afectados
        self.modified_contacts = {}
        
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
            logo_path = resource_path("logo.png")
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
        
        # Botón para agregar contacto manualmente
        ttk.Button(import_buttons_frame, text="Agregar contacto manualmente", 
                  command=self.add_contact_manually, style="success.TButton").pack(side=tk.LEFT, padx=5)
        
        # Añadir campo de búsqueda
        search_frame = ttk.Frame(import_buttons_frame)
        search_frame.pack(side=tk.RIGHT, padx=10, fill=tk.X, expand=True)
        
        ttk.Label(search_frame, text="Buscar:").pack(side=tk.LEFT, padx=(10, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=25)
        self.search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(search_frame, text="Buscar", 
                  command=self.search_contacts, style="info.TButton").pack(side=tk.LEFT, padx=5)
        
        # Botón para resetear/limpiar búsqueda
        ttk.Button(search_frame, text="Limpiar", 
                  command=self.reset_search, style="secondary.TButton").pack(side=tk.LEFT, padx=5)
        
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
        ttk.Button(action_frame, text="Eliminar Contacto", 
                  command=self.delete_contacts, style="danger.TButton").pack(side=tk.LEFT, padx=5)
        
        # Contador de clientes
        self.client_count_label = ttk.Label(action_frame, text="Total Clientes: 0", font=("Helvetica", 9, "bold"))
        self.client_count_label.pack(side=tk.LEFT, padx=15)
        
        # Contador de comerciales
        self.commercial_count_label = ttk.Label(action_frame, text="Total Comerciales: 0", font=("Helvetica", 9, "bold"))
        self.commercial_count_label.pack(side=tk.LEFT, padx=15)
        
        ttk.Button(action_frame, text="Reclasificar Contacto", 
                  command=self.reclassify_contact, style="warning.TButton").pack(side=tk.RIGHT, padx=5)
        
        # Footer con información de copyright
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=5)
        
        # Botón para activar/desactivar modo de depuración (oculto en una esquina)
        debug_frame = ttk.Frame(footer_frame)
        debug_frame.pack(side=tk.RIGHT, padx=10)
        
        self.debug_var = tk.BooleanVar(value=False)
        self.debug_button = ttk.Checkbutton(debug_frame, text="Modo depuración", 
                                          variable=self.debug_var, 
                                          command=self.toggle_debug_mode)
        self.debug_button.pack(side=tk.RIGHT, padx=5)
        
        footer_text = "Diseñado con ❤️ por franHR\nCopyright © 2025 pcprogramacion.es"
        footer_label = ttk.Label(footer_frame, text=footer_text, justify=tk.CENTER, 
                                 font=("Helvetica", 9))
        footer_label.pack(pady=10)
        
        # Línea separadora para el footer
        ttk.Separator(main_frame).pack(fill=tk.X, before=footer_frame)
        
        # Cargar datos iniciales
        self.update_lists()
    
    def toggle_debug_mode(self):
        """Activa o desactiva el modo de depuración."""
        debug_enabled = self.debug_var.get()
        self.db_manager.set_debug_mode(debug_enabled)
        status = "ACTIVADO" if debug_enabled else "DESACTIVADO"
        print(f"Modo de depuración {status}")
        if debug_enabled:
            messagebox.showinfo("Modo depuración", 
                              "El modo de depuración ha sido activado. Se mostrarán detalles adicionales en la consola.")
        else:
            messagebox.showinfo("Modo depuración", 
                              "El modo de depuración ha sido desactivado.")
    
    def search_contacts(self):
        """Busca contactos en todas las categorías basándose en el texto de búsqueda."""
        search_text = self.search_var.get().strip().lower()
        if not search_text:
            # Si el campo está vacío, mostrar todos los contactos
            messagebox.showinfo("Búsqueda", "Por favor, ingrese un texto para buscar.")
            self.update_lists()
            return
        
        # Limpiar treeviews actuales
        for item in self.clients_tree.get_children():
            self.clients_tree.delete(item)
        
        for item in self.commercial_tree.get_children():
            self.commercial_tree.delete(item)
        
        for item in self.invalid_tree.get_children():
            self.invalid_tree.delete(item)
        
        # Buscar en la tabla de clientes
        cursor = self.db_manager.conn.cursor()
        cursor.execute("""
            SELECT name, email, imported_date, client_code, address, postal_code, town, city, additional_info 
            FROM clients 
            WHERE 
                LOWER(name) LIKE ? OR 
                LOWER(email) LIKE ? OR 
                LOWER(client_code) LIKE ? OR 
                LOWER(address) LIKE ? OR 
                LOWER(postal_code) LIKE ? OR 
                LOWER(town) LIKE ? OR 
                LOWER(city) LIKE ? OR 
                LOWER(additional_info) LIKE ?
        """, (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", 
              f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"))
        
        clients = cursor.fetchall()
        for client in clients:
            client_id = f"{client[0]}_{client[1]}"  # Usar combinación de nombre y email como ID
            if client_id in self.new_contacts:
                self.clients_tree.insert("", tk.END, values=client, tags=('new_contact',))
            else:
                self.clients_tree.insert("", tk.END, values=client)
        
        # Buscar en la tabla de contactos comerciales
        cursor.execute("""
            SELECT name, email, company, imported_date, client_code, address, postal_code, town, city, additional_info 
            FROM commercial_contacts 
            WHERE 
                LOWER(name) LIKE ? OR 
                LOWER(email) LIKE ? OR 
                LOWER(company) LIKE ? OR
                LOWER(client_code) LIKE ? OR 
                LOWER(address) LIKE ? OR 
                LOWER(postal_code) LIKE ? OR 
                LOWER(town) LIKE ? OR 
                LOWER(city) LIKE ? OR 
                LOWER(additional_info) LIKE ?
        """, (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", 
              f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"))
        
        commercials = cursor.fetchall()
        for commercial in commercials:
            commercial_id = f"{commercial[0]}_{commercial[1]}"  # Usar combinación de nombre y email como ID
            if commercial_id in self.new_contacts:
                self.commercial_tree.insert("", tk.END, values=commercial, tags=('new_contact',))
            else:
                self.commercial_tree.insert("", tk.END, values=commercial)
        
        # Buscar en la tabla de emails no válidos
        cursor.execute("""
            SELECT email, name, imported_date, reason 
            FROM invalid_emails 
            WHERE 
                LOWER(email) LIKE ? OR 
                LOWER(name) LIKE ? OR 
                LOWER(reason) LIKE ?
        """, (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"))
        
        invalids = cursor.fetchall()
        for invalid in invalids:
            self.invalid_tree.insert("", tk.END, values=invalid)
        
        # Mostrar un mensaje con los resultados
        total = len(clients) + len(commercials) + len(invalids)
        messagebox.showinfo("Resultados de búsqueda", 
                          f"Se encontraron {total} contactos que coinciden con '{search_text}'.\n"
                          f"• Clientes: {len(clients)}\n"
                          f"• Contactos comerciales: {len(commercials)}\n"
                          f"• Emails no válidos: {len(invalids)}")
    
    def setup_clients_treeview(self):
        """Configura el treeview para clientes regulares."""
        columns = ("name", "email", "date", "client_code", "address", "postal_code", "town", "city", "additional_info")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.clients_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.clients_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                       selectmode="extended")
        
        self.clients_tree.heading("name", text="Nombre")
        self.clients_tree.heading("email", text="Email")
        self.clients_tree.heading("date", text="Fecha Importación")
        self.clients_tree.heading("client_code", text="Código de Cliente")
        self.clients_tree.heading("address", text="Dirección")
        self.clients_tree.heading("postal_code", text="Código Postal")
        self.clients_tree.heading("town", text="Población")
        self.clients_tree.heading("city", text="Ciudad")
        self.clients_tree.heading("additional_info", text="Información Adicional")
        
        self.clients_tree.column("name", width=200)
        self.clients_tree.column("email", width=250)
        self.clients_tree.column("date", width=150)
        self.clients_tree.column("client_code", width=100)
        self.clients_tree.column("address", width=200)
        self.clients_tree.column("postal_code", width=100)
        self.clients_tree.column("town", width=150)
        self.clients_tree.column("city", width=150)
        self.clients_tree.column("additional_info", width=200)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.clients_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.clients_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.clients_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.clients_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar el color verde para los nuevos contactos
        self.clients_tree.tag_configure('new_contact', background='#c8ffb3')
        # Configurar el color verde para las celdas modificadas
        self.clients_tree.tag_configure('modified_cell', background='#a5e88f')
    
    def setup_commercial_treeview(self):
        """Configura el treeview para contactos comerciales."""
        columns = ("name", "email", "company", "date", "client_code", "address", "postal_code", "town", "city", "additional_info")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.commercial_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.commercial_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                          selectmode="extended")
        
        self.commercial_tree.heading("name", text="Nombre")
        self.commercial_tree.heading("email", text="Email")
        self.commercial_tree.heading("company", text="Empresa")
        self.commercial_tree.heading("date", text="Fecha Importación")
        self.commercial_tree.heading("client_code", text="Código de Cliente")
        self.commercial_tree.heading("address", text="Dirección")
        self.commercial_tree.heading("postal_code", text="Código Postal")
        self.commercial_tree.heading("town", text="Población")
        self.commercial_tree.heading("city", text="Ciudad")
        self.commercial_tree.heading("additional_info", text="Información Adicional")
        
        self.commercial_tree.column("name", width=180)
        self.commercial_tree.column("email", width=220)
        self.commercial_tree.column("company", width=180)
        self.commercial_tree.column("date", width=120)
        self.commercial_tree.column("client_code", width=100)
        self.commercial_tree.column("address", width=200)
        self.commercial_tree.column("postal_code", width=100)
        self.commercial_tree.column("town", width=150)
        self.commercial_tree.column("city", width=150)
        self.commercial_tree.column("additional_info", width=200)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.commercial_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.commercial_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.commercial_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.commercial_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar el color verde para los nuevos contactos
        self.commercial_tree.tag_configure('new_contact', background='#c8ffb3')
        # Configurar el color verde para las celdas modificadas
        self.commercial_tree.tag_configure('modified_cell', background='#a5e88f')
    
    def setup_invalid_treeview(self):
        """Configura el treeview para emails no válidos."""
        columns = ("email", "name", "date", "reason")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.invalid_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.invalid_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                       selectmode="extended")
        
        self.invalid_tree.heading("email", text="Email")
        self.invalid_tree.heading("name", text="Nombre")
        self.invalid_tree.heading("date", text="Fecha Importación")
        self.invalid_tree.heading("reason", text="Motivo")
        
        self.invalid_tree.column("email", width=250)
        self.invalid_tree.column("name", width=200)
        self.invalid_tree.column("date", width=150)
        self.invalid_tree.column("reason", width=150)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.invalid_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.invalid_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.invalid_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.invalid_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
    
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
        regular_clients = []
        new_clients = []
        
        # Separar clientes regulares y nuevos
        for client in clients:
            client_id = f"{client[0]}_{client[1]}"  # Usar combinación de nombre y email como ID
            if client_id in self.new_contacts:
                new_clients.append(client)
            else:
                regular_clients.append(client)
        
        # Mostrar clientes regulares
        for client in regular_clients:
            email = client[1]  # El email está en la posición 1
            
            # Comprobar si el cliente está en la lista de contactos modificados
            if email in self.modified_contacts:
                # Si está modificado, insertar con la etiqueta de modificado
                item_id = self.clients_tree.insert("", tk.END, values=client, tags=('modified_cell',))
            else:
                # Si no está modificado, insertar normalmente
                item_id = self.clients_tree.insert("", tk.END, values=client)
        
        # Mostrar clientes nuevos al final con tag especial
        for client in new_clients:
            self.clients_tree.insert("", tk.END, values=client, tags=('new_contact',))
        
        # Actualizar contador de clientes
        total_clients = len(regular_clients) + len(new_clients)
        self.client_count_label.config(text=f"Total Clientes: {total_clients}")
        
        # Obtener y mostrar contactos comerciales
        commercials = self.db_manager.get_all_commercial_contacts()
        regular_commercials = []
        new_commercials = []
        
        # Separar contactos comerciales regulares y nuevos
        for commercial in commercials:
            commercial_id = f"{commercial[0]}_{commercial[1]}"  # Usar combinación de nombre y email como ID
            if commercial_id in self.new_contacts:
                new_commercials.append(commercial)
            else:
                regular_commercials.append(commercial)
        
        # Mostrar contactos comerciales regulares
        for commercial in regular_commercials:
            email = commercial[1]  # El email está en la posición 1
            
            # Comprobar si el contacto comercial está en la lista de contactos modificados
            if email in self.modified_contacts:
                # Si está modificado, insertar con la etiqueta de modificado
                item_id = self.commercial_tree.insert("", tk.END, values=commercial, tags=('modified_cell',))
            else:
                # Si no está modificado, insertar normalmente
                item_id = self.commercial_tree.insert("", tk.END, values=commercial)
        
        # Mostrar contactos comerciales nuevos al final con tag especial
        for commercial in new_commercials:
            self.commercial_tree.insert("", tk.END, values=commercial, tags=('new_contact',))
        
        # Actualizar contador de contactos comerciales
        total_commercials = len(regular_commercials) + len(new_commercials)
        self.commercial_count_label.config(text=f"Total Comerciales: {total_commercials}")
        
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
            
        # Variable para la ventana de progreso
        progress_window = None
        
        # Variable para la ventana de progreso
        progress_window = None
        
        try:
            # Crear ventana de progreso
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Importando CSV")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            self.center_window(progress_window)
            
            # Configurar como ventana modal
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # Añadir mensaje y barra de progreso
            ttk.Label(
                progress_window, 
                text="Importando datos del archivo CSV...\nPor favor espere mientras se procesa el archivo.",
                font=("Segoe UI", 10),
                wraplength=380
            ).pack(pady=15)
            
            progress_bar = ttk.Progressbar(
                progress_window, 
                mode="indeterminate", 
                length=350,
                style="primary.Horizontal.TProgressbar"
            )
            progress_bar.pack(pady=10)
            progress_bar.start(10)
            
            # Actualizar la ventana para asegurar que se muestre
            progress_window.update()
            
            # Reiniciar la opción de omitir todos
            self.skip_all_duplicates = False
            
            # Probar diferentes codificaciones
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
            
            for encoding in encodings:
                try:
                    # Intentar abrir el archivo con la codificación actual
                    with open(filepath, 'r', encoding=encoding) as test_file:
                        test_file.readline()  # Leer una línea para verificar
                    
                    # Si llegamos aquí, la codificación funciona
                    break
                except UnicodeDecodeError:
                    if encoding == encodings[-1]:
                        # Si es la última codificación y falla, mostrar error
                        messagebox.showerror("Error", f"No se pudo determinar la codificación del archivo. Prueba guardarlo como UTF-8.")
                        return
                    continue  # Probar con la siguiente codificación
            
            # Primero leer el archivo para determinar las columnas
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')  # Usar ; como delimitador por defecto
                headers = next(csv_reader, None)  # Leer encabezados
            
            if not headers:
                messagebox.showerror("Error", "El archivo CSV no tiene encabezados o está vacío.")
                return
            
            # Crear ventana para seleccionar las columnas
            select_columns_window = tk.Toplevel(self.root)
            select_columns_window.title("Seleccionar columnas")
            # Aumentar el tamaño de la ventana de selección de columnas (era 800x600)
            select_columns_window.geometry("1000x700")
            select_columns_window.resizable(True, True)
            select_columns_window.transient(self.root)
            select_columns_window.grab_set()
            select_columns_window.configure(bg="#f8f4ef")
            
            frame = ttk.Frame(select_columns_window, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Información
            ttk.Label(frame, text="Seleccione las columnas que contienen el email y el nombre:", 
                     font=("Helvetica", 10, "bold")).pack(pady=(5, 10))
            
            # Variables para almacenar los índices seleccionados
            # Obtener los valores guardados de la configuración
            saved_email_index = self.db_manager.get_config("email_index", "0")
            saved_name_index = self.db_manager.get_config("name_index", "0")
            saved_client_code_index = self.db_manager.get_config("client_code_index", "-1")
            saved_address_index = self.db_manager.get_config("address_index", "-1")
            saved_postal_code_index = self.db_manager.get_config("postal_code_index", "-1")
            saved_town_index = self.db_manager.get_config("town_index", "-1")
            saved_city_index = self.db_manager.get_config("city_index", "-1")
            
            email_index_var = tk.StringVar(value=saved_email_index)
            name_index_var = tk.StringVar(value=saved_name_index)
            # Additional fields for new columns
            client_code_index_var = tk.StringVar(value=saved_client_code_index)
            address_index_var = tk.StringVar(value=saved_address_index)
            postal_code_index_var = tk.StringVar(value=saved_postal_code_index)
            town_index_var = tk.StringVar(value=saved_town_index)
            city_index_var = tk.StringVar(value=saved_city_index)
            
            # Crear marco para la tabla de vista previa
            preview_frame = ttk.LabelFrame(frame, text="Vista previa del CSV")
            preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # Crear un canvas y un frame dentro para la tabla
            canvas = tk.Canvas(preview_frame, bg="#f8f4ef")
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=canvas.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            canvas.configure(yscrollcommand=scrollbar.set)
            
            preview_table_frame = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=preview_table_frame, anchor=tk.NW)
            
            # Mostrar las primeras filas del CSV para ayudar en la selección
            num_preview_rows = 5  # Número de filas para vista previa
            
            # Reiniciar el archivo para leer los datos de vista previa
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')
                preview_rows = [next(csv_reader, None)]  # Encabezados
                
                for _ in range(num_preview_rows):
                    try:
                        preview_rows.append(next(csv_reader, None))
                    except StopIteration:
                        break
            
            # Mostrar los datos en una tabla
            max_columns = max(len(row) if row else 0 for row in preview_rows)
            
            # Labels para los índices de columna
            for col in range(max_columns):
                ttk.Label(preview_table_frame, text=f"Col {col}", font=("Helvetica", 9, "bold"), 
                         borderwidth=1, relief="solid", width=15, anchor=tk.CENTER).grid(row=0, column=col, padx=2, pady=2)
            
            # Datos de vista previa
            for row_idx, row in enumerate(preview_rows):
                if row:
                    for col_idx, value in enumerate(row):
                        if col_idx < max_columns:  # Asegurarse de no exceder el número máximo de columnas
                            ttk.Label(preview_table_frame, text=value[:20] + ('...' if len(value) > 20 else ''), 
                                     borderwidth=1, relief="solid", width=15).grid(row=row_idx+1, column=col_idx, padx=2, pady=2)
            
            # Actualizar tamaño del canvas después de agregar widgets
            preview_table_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))
            
            # Selección de columnas
            selection_frame = ttk.Frame(frame)
            selection_frame.pack(fill=tk.X, pady=10)
            
            # Crear selectores para cada campo
            email_selector = self.create_column_selector(selection_frame, email_index_var, "Email (obligatorio):", headers)
            name_selector = self.create_column_selector(selection_frame, name_index_var, "Nombre (obligatorio):", headers)
            
            # Campos adicionales
            client_code_selector = self.create_column_selector(selection_frame, client_code_index_var, "Código Cliente:", headers)
            address_selector = self.create_column_selector(selection_frame, address_index_var, "Dirección:", headers)
            postal_code_selector = self.create_column_selector(selection_frame, postal_code_index_var, "Código Postal:", headers)
            town_selector = self.create_column_selector(selection_frame, town_index_var, "Población:", headers)
            city_selector = self.create_column_selector(selection_frame, city_index_var, "Ciudad:", headers)
            
            # Variable para almacenar la respuesta
            result = {"proceed": False, "email_index": 0, "name_index": 0, "client_code_index": -1, "address_index": -1, "postal_code_index": -1, "town_index": -1, "city_index": -1}
            
            # Función para procesar la selección
            def on_submit():
                result["proceed"] = True
                result["email_index"] = email_index_var.get()
                result["name_index"] = name_index_var.get()
                result["client_code_index"] = client_code_index_var.get()
                result["address_index"] = address_index_var.get()
                result["postal_code_index"] = postal_code_index_var.get()
                result["town_index"] = town_index_var.get()
                result["city_index"] = city_index_var.get()
                select_columns_window.destroy()
            
            # Botones
            buttons_frame = ttk.Frame(frame)
            buttons_frame.pack(fill=tk.X, pady=10)
            
            ttk.Button(buttons_frame, text="Cancelar", 
                      command=select_columns_window.destroy).pack(side=tk.LEFT, padx=10)
            
            ttk.Button(buttons_frame, text="Importar", 
                      command=on_submit, style="success.TButton").pack(side=tk.RIGHT, padx=10)
            
            # Centrar la ventana
            self.center_window(select_columns_window)
            
            # Esperar a que se cierre la ventana
            self.root.wait_window(select_columns_window)
            
            if not result["proceed"]:
                return
            
            # Guardar la configuración para futuras importaciones
            self.db_manager.save_config("email_index", str(result["email_index"]))
            self.db_manager.save_config("name_index", str(result["name_index"]))
            self.db_manager.save_config("client_code_index", str(result["client_code_index"]))
            self.db_manager.save_config("address_index", str(result["address_index"]))
            self.db_manager.save_config("postal_code_index", str(result["postal_code_index"]))
            self.db_manager.save_config("town_index", str(result["town_index"]))
            self.db_manager.save_config("city_index", str(result["city_index"]))
            
            # Asegurarnos de que la conexión a la base de datos está activa
            if self.db_manager.conn is None or not hasattr(self.db_manager, 'conn'):
                self.db_manager.create_database()
            
            # Limpiar la lista de nuevos contactos antes de importar
            self.new_contacts = []
            
            # Convertir índices a enteros
            email_index = int(result["email_index"])
            name_index = int(result["name_index"])
            client_code_index = int(result["client_code_index"])
            address_index = int(result["address_index"])
            postal_code_index = int(result["postal_code_index"])
            town_index = int(result["town_index"])
            city_index = int(result["city_index"])
            
            # SOLUCIÓN: Leer todo el CSV y luego procesar por email para garantizar coherencia
            # Primero leemos todo el CSV y lo organizamos por email
            csv_data = {}
            
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')
                # Eliminar la siguiente línea para no saltar encabezados
                # next(csv_reader, None)  # Saltar encabezados
                
                # Contadores para depuración
                skipped_insufficient_columns = 0
                skipped_empty_email = 0
                
                for row in csv_reader:
                    # Verificar que la fila tenga suficientes columnas
                    max_required_idx = max(email_index, name_index, 
                                       client_code_index if client_code_index > 0 else 0,
                                       address_index if address_index > 0 else 0,
                                       postal_code_index if postal_code_index > 0 else 0,
                                       town_index if town_index > 0 else 0,
                                       city_index if city_index > 0 else 0)
                    
                    if len(row) <= max_required_idx:
                        skipped_insufficient_columns += 1
                        if self.db_manager.debug_mode:
                            print(f"DEBUG: Saltando fila con columnas insuficientes. Tiene {len(row)} columnas, se necesitan {max_required_idx+1}.")
                            if len(row) > 0:
                                print(f"Contenido de la fila: {row}")
                        continue  # Saltar filas que no tienen suficientes columnas
                    
                    email = row[email_index].strip().lower() if email_index < len(row) else ""
                    
                    if not email:
                        skipped_empty_email += 1
                        if self.db_manager.debug_mode:
                            print(f"DEBUG: Saltando fila sin email. Contenido: {row}")
                        continue  # Saltar filas sin email
                    
                    # Guardar los datos de esta fila organizados por email
                    csv_data[email] = {
                        'name': row[name_index].strip() if name_index < len(row) else "",
                        'client_code': row[client_code_index].strip() if client_code_index >= 0 and client_code_index < len(row) else "",
                        'address': row[address_index].strip() if address_index >= 0 and address_index < len(row) else "",
                        'postal_code': row[postal_code_index].strip() if postal_code_index >= 0 and postal_code_index < len(row) else "",
                        'town': row[town_index].strip() if town_index >= 0 and town_index < len(row) else "",
                        'city': row[city_index].strip() if city_index >= 0 and city_index < len(row) else ""
                    }
            
            # Contadores para el informe final
            new_clients = 0
            updated_clients = 0
            unchanged_clients = 0
            duplicate_contacts = 0
            clients_with_changes = []
            
            # Ahora procesamos cada email asegurándonos que la asociación es correcta
            for email, data in csv_data.items():
                name = data['name']
                client_code = data['client_code']
                address = data['address']
                postal_code = data['postal_code']
                town = data['town']
                city = data['city']
                
                # Verificar si el email ya existe en alguna tabla
                existing_category = self.check_existing_email(email)
                
                if existing_category:
                    # Comprobar si existe en la base de datos y verificar si hay cambios
                    if existing_category == "client":
                        # Si es un cliente existente, intentar actualizar o detectar cambios
                        result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                        
                        # Procesar el resultado
                        if "new" in result and result["new"]:
                            new_clients += 1
                            self.new_contacts.append(f"{name}_{email}")
                        
                        elif "updated" in result:
                            if result["updated"] and result["changes"]:
                                # Solo mostrar diálogo si hay cambios reales
                                # Preguntar al usuario si desea aplicar los cambios
                                changes_message = f"Se han detectado cambios en el cliente {name} <{email}>:\n\n"
                                for change in result["changes"]:
                                    changes_message += f"• {change}\n"
                                changes_message += "\n¿Desea aplicar estos cambios?"
                                
                                # Añadir opción para activar depuración si hay muchos cambios
                                if len(result["changes"]) > 3:
                                    changes_message += "\n\nNota: Si cree que estos cambios no son correctos, puede activar el modo de depuración en la parte inferior derecha de la ventana para obtener más información."
                                
                                if messagebox.askyesno("Cambios detectados", changes_message):
                                    # Si el usuario acepta, actualizar el contador
                                    updated_clients += 1
                                    # Guardar los detalles de los cambios
                                    clients_with_changes.append({
                                        "name": name,
                                        "email": email,
                                        "changes": result["changes"]
                                    })
                                    
                                    # Guardar los campos modificados para colorearlos después
                                    modified_fields = []
                                    for change in result["changes"]:
                                        parts = change.split(": ")
                                        if len(parts) >= 2:
                                            field_name = parts[0]
                                            # Mapear los nombres de campo en español a los nombres de columna en inglés
                                            field_mapping = {
                                                "nombre": "name",
                                                "código cliente": "client_code",
                                                "dirección": "address",
                                                "código postal": "postal_code",
                                                "población": "town",
                                                "ciudad": "city"
                                            }
                                            if field_name in field_mapping:
                                                modified_fields.append(field_mapping[field_name])
                                    
                                    # Guardar en el diccionario de contactos modificados
                                    self.modified_contacts[email] = modified_fields
                                else:
                                    # Si el usuario rechaza, revertir los cambios en la base de datos
                                    cursor = self.db_manager.conn.cursor()
                                    cursor.execute("SELECT id FROM clients WHERE email = ?", (email,))
                                    client_id = cursor.fetchone()[0]
                                    
                                    # Buscar los valores originales
                                    for change in result["changes"]:
                                        parts = change.split(": ")
                                        if len(parts) >= 2:
                                            field = parts[0]
                                            values_part = parts[1]
                                            value_parts = values_part.split(" -> ")
                                            if len(value_parts) >= 1:
                                                original = value_parts[0].strip("'")
                                                
                                                # Actualizar los campos que han cambiado
                                                if "nombre" in field:
                                                    cursor.execute("UPDATE clients SET name = ? WHERE id = ?", (original, client_id))
                                                elif "código cliente" in field:
                                                    cursor.execute("UPDATE clients SET client_code = ? WHERE id = ?", (original, client_id))
                                                elif "dirección" in field:
                                                    cursor.execute("UPDATE clients SET address = ? WHERE id = ?", (original, client_id))
                                                elif "código postal" in field:
                                                    cursor.execute("UPDATE clients SET postal_code = ? WHERE id = ?", (original, client_id))
                                                elif "población" in field:
                                                    cursor.execute("UPDATE clients SET town = ? WHERE id = ?", (original, client_id))
                                                elif "ciudad" in field:
                                                    cursor.execute("UPDATE clients SET city = ? WHERE id = ?", (original, client_id))
                                    
                                    self.db_manager.conn.commit()
                                    unchanged_clients += 1
                            else:
                                # Si no hay cambios, incrementar el contador de sin cambios
                                unchanged_clients += 1
                        elif "error" in result:
                            print(f"Error al procesar cliente {name}, {email}: {result['error']}")
                    elif existing_category == "commercial":
                        # Si es un contacto comercial existente, utilizar el nuevo método con detección de cambios
                        result = self.db_manager.add_commercial_contact_with_changes(
                            name, email, "", client_code, address, postal_code, town, city
                        )
                        
                        if "new" in result and result["new"]:
                            # No debería ser un nuevo contacto si ya existe
                            pass
                        elif "updated" in result:
                            if result["updated"] and result["changes"]:
                                # Filtrar solo los cambios reales en los datos
                                real_changes = []
                                for change in result["changes"]:
                                    # Ignorar cambios de categoría
                                    if not any(cat in change.lower() for cat in ["categoría", "category"]):
                                        real_changes.append(change)
                                
                                # Verificar si el contacto acaba de ser reclasificado recientemente
                                cursor = self.db_manager.conn.cursor()
                                
                                # Solo considerar como "recién movido" si fue añadido hace menos de 10 minutos
                                # Lo que generalmente indica que fue parte de la misma operación
                                cursor.execute("SELECT imported_date FROM commercial_contacts WHERE email = ?", (email,))
                                commercial_date = cursor.fetchone()
                                contact_just_moved = False
                                
                                if commercial_date:
                                    last_date = commercial_date[0]
                                    now = datetime.now()
                                    
                                    if isinstance(last_date, str):
                                        try:
                                            last_date = datetime.strptime(last_date, "%Y-%m-%d %H:%M:%S")
                                            # Si fue agregado en los últimos 10 minutos, probablemente es 
                                            # parte de la misma operación de reclasificación
                                            if (now - last_date).total_seconds() < 600:  # 10 minutos
                                                # Buscar si existía antes en otra tabla (cliente)
                                                cursor.execute("SELECT COUNT(*) FROM clients WHERE email = ?", (email,))
                                                client_count = cursor.fetchone()[0]
                                                if client_count > 0:
                                                    # Si existe en ambas tablas, es una reclasificación en progreso
                                                    contact_just_moved = True
                                        except ValueError:
                                            pass
                                
                                # Si hay cambios importantes en el contacto (no solo cambio de categoría)
                                if real_changes:
                                    # Preguntar al usuario si desea aplicar los cambios
                                    changes_message = f"Se han detectado cambios en el contacto comercial {name} <{email}>:\n\n"
                                    for change in real_changes:
                                        changes_message += f"• {change}\n"
                                    changes_message += "\n¿Desea aplicar estos cambios?"
                                    
                                    # Añadir opción para activar depuración si hay muchos cambios
                                    if len(real_changes) > 3:
                                        changes_message += "\n\nNota: Si cree que estos cambios no son correctos, puede activar el modo de depuración en la parte inferior derecha de la ventana para obtener más información."
                                    
                                    if messagebox.askyesno("Cambios detectados", changes_message):
                                        # Si el usuario acepta, actualizar el contador
                                        updated_clients += 1
                                        # Guardar los detalles de los cambios
                                        clients_with_changes.append({
                                            "name": name,
                                            "email": email,
                                            "changes": real_changes
                                        })
                                        
                                        # Guardar los campos modificados para colorearlos después
                                        modified_fields = []
                                        for change in real_changes:
                                            parts = change.split(": ")
                                            if len(parts) >= 2:
                                                field_name = parts[0]
                                                # Mapear los nombres de campo en español a los nombres de columna en inglés
                                                field_mapping = {
                                                    "nombre": "name",
                                                    "empresa": "company",
                                                    "código cliente": "client_code",
                                                    "dirección": "address",
                                                    "código postal": "postal_code",
                                                    "población": "town",
                                                    "ciudad": "city",
                                                    "información adicional": "additional_info"
                                                }
                                                if field_name in field_mapping:
                                                    modified_fields.append(field_mapping[field_name])
                                        
                                        # Guardar en el diccionario de contactos modificados
                                        self.modified_contacts[email] = modified_fields
                                else:
                                    # No hay cambios reales en los datos
                                    unchanged_clients += 1
                            else:
                                # No hay cambios según la BD
                                unchanged_clients += 1
                        elif "error" in result:
                            print(f"Error al procesar contacto comercial {name}, {email}: {result['error']}")
                    elif existing_category == "invalid":
                        # Si es un email no válido, simplemente omitirlo
                        duplicate_contacts += 1
                        continue
                    elif not self.skip_all_duplicates:
                        # Solo mostrar diálogo para contactos que existen en OTRA categoría (no como cliente)
                        action = self.ask_duplicate_action(email, name, existing_category)
                        
                        if action == "skip":
                            # Omitir este email
                            continue
                        elif action == "skip_all":
                            # Omitir todos los duplicados restantes
                            self.skip_all_duplicates = True
                            continue
                        elif action == "client":
                            # Reclasificar a cliente
                            self.delete_from_all_categories(email)
                            result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                            if "new" in result and result["new"]:
                                # Agregar a la lista de nuevos contactos
                                self.new_contacts.append(f"{name}_{email}")
                            continue
                    else:
                        # Si omitir todos está marcado, simplemente saltamos este email
                        continue
                else:
                    # Si llegamos aquí, es un nuevo contacto (no existe en la base de datos)
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    if "new" in result and result["new"]:
                        new_clients += 1
                        # Agregar a la lista de nuevos contactos
                        self.new_contacts.append(f"{name}_{email}")
            
            self.update_lists()
            
            # Preparar mensaje detallado
            detailed_message = f"Se han añadido {new_clients} nuevos clientes.\n"\
                              f"Se han actualizado {updated_clients} clientes existentes.\n"\
                              f"Se han omitido {unchanged_clients} clientes sin cambios.\n"\
                              f"Se han omitido {duplicate_contacts} contactos duplicados."
            
            # Añadir información de depuración sobre filas saltadas
            if self.db_manager.debug_mode:
                total_rows = skipped_insufficient_columns + skipped_empty_email + len(csv_data)
                detailed_message += f"\n\nInformación de depuración:\n"\
                                   f"- Total de filas en CSV: {total_rows}\n"\
                                   f"- Filas procesadas correctamente: {len(csv_data)}\n"\
                                   f"- Filas saltadas por columnas insuficientes: {skipped_insufficient_columns}\n"\
                                   f"- Filas saltadas por email vacío: {skipped_empty_email}\n"\
                                   f"- Duplicados (mismo email): {total_rows - skipped_insufficient_columns - skipped_empty_email - len(csv_data)}"
            
            # Si hay clientes actualizados, mostrar los detalles de los cambios
            if updated_clients > 0 and clients_with_changes:
                changes_details = "\n\nDetalles de los cambios:\n"
                for client in clients_with_changes:
                    changes_details += f"\n- {client['name']} ({client['email']}):\n"
                    for change in client['changes']:
                        changes_details += f"  • {change}\n"
                
                # Si el mensaje es muy largo, preguntar si quiere ver los detalles
                if len(changes_details) > 500:
                    show_details = messagebox.askyesno(
                        "Cambios Detallados",
                        f"{detailed_message}\n\n¿Desea ver los detalles de todos los cambios?"
                    )
                    if show_details:
                        # Mostrar detalles en una ventana de texto desplazable
                        details_window = tk.Toplevel(self.root)
                        details_window.title("Detalles de Cambios")
                        details_window.geometry("700x500")
                        details_window.transient(self.root)
                        details_window.grab_set()
                        
                        # Frame principal
                        frame = ttk.Frame(details_window, padding=10)
                        frame.pack(fill=tk.BOTH, expand=True)
                        
                        # Título
                        ttk.Label(
                            frame, 
                            text="Detalles de contactos actualizados", 
                            font=("Helvetica", 12, "bold")
                        ).pack(anchor=tk.W, pady=(0, 10))
                        
                        # Área de texto con scroll
                        text_frame = ttk.Frame(frame)
                        text_frame.pack(fill=tk.BOTH, expand=True)
                        
                        text = tk.Text(text_frame, wrap=tk.WORD)
                        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                        
                        scrollbar = ttk.Scrollbar(text_frame, command=text.yview)
                        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                        
                        text.configure(yscrollcommand=scrollbar.set)
                        text.insert(tk.END, changes_details)
                        text.configure(state="disabled")  # Solo lectura
                        
                        # Botón para cerrar
                        ttk.Button(
                            frame, 
                            text="Cerrar", 
                            command=details_window.destroy
                        ).pack(pady=10)
                        
                        # Centrar la ventana relativa a la ventana principal
                        self.center_window(details_window)
                else:
                    # Si el mensaje no es muy largo, mostrar todo junto
                    messagebox.showinfo(
                        "Importación Completada", 
                        f"{detailed_message}\n{changes_details}"
                    )
            else:
                # Si no hay cambios, mostrar solo el mensaje básico
                messagebox.showinfo("Importación Completada", detailed_message)
            
            # Limpiar variables que controlan la importación para próximos usos
            self.skip_all_duplicates = False
            
            # No limpiamos self.new_contacts para que se muestren resaltados hasta la próxima importación
            # En cambio, limpiamos los contactos modificados ya que esos sí tienen el efecto permanente en la BD
            # y ya han sido vistos en el informe que se acaba de mostrar
            self.modified_contacts = {}
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar el archivo: {str(e)}")
            
        finally:
            # Cerrar la ventana de progreso SIEMPRE, incluso si hay excepciones
            if 'progress_window' in locals() and progress_window and progress_window.winfo_exists():
                progress_window.destroy()
    
    def add_contact_manually(self):
        """Permite agregar un contacto manualmente."""
        # Crear ventana para añadir contacto
        add_window = tk.Toplevel(self.root)
        add_window.title("Agregar contacto manualmente")
        add_window.geometry("500x550")
        add_window.resizable(False, False)
        add_window.transient(self.root)
        add_window.grab_set()
        add_window.configure(bg="#f8f4ef")
        
        frame = ttk.Frame(add_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Ingrese los datos del nuevo contacto:", 
                 font=("Helvetica", 10, "bold")).pack(pady=(5, 15))
        
        # Campos para nombre y email
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(name_frame, text="Nombre:").pack(side=tk.LEFT, padx=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        email_frame = ttk.Frame(frame)
        email_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(email_frame, text="Email:").pack(side=tk.LEFT, padx=5)
        email_var = tk.StringVar()
        email_entry = ttk.Entry(email_frame, textvariable=email_var, width=30)
        email_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Nuevos campos
        # Código de cliente
        client_code_frame = ttk.Frame(frame)
        client_code_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(client_code_frame, text="Código de Cliente:").pack(side=tk.LEFT, padx=5)
        client_code_var = tk.StringVar()
        client_code_entry = ttk.Entry(client_code_frame, textvariable=client_code_var, width=30)
        client_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Dirección
        address_frame = ttk.Frame(frame)
        address_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(address_frame, text="Dirección:").pack(side=tk.LEFT, padx=5)
        address_var = tk.StringVar()
        address_entry = ttk.Entry(address_frame, textvariable=address_var, width=30)
        address_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Código postal
        postal_code_frame = ttk.Frame(frame)
        postal_code_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(postal_code_frame, text="Código Postal:").pack(side=tk.LEFT, padx=5)
        postal_code_var = tk.StringVar()
        postal_code_entry = ttk.Entry(postal_code_frame, textvariable=postal_code_var, width=30)
        postal_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Población
        town_frame = ttk.Frame(frame)
        town_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(town_frame, text="Población:").pack(side=tk.LEFT, padx=5)
        town_var = tk.StringVar()
        town_entry = ttk.Entry(town_frame, textvariable=town_var, width=30)
        town_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Ciudad
        city_frame = ttk.Frame(frame)
        city_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(city_frame, text="Ciudad:").pack(side=tk.LEFT, padx=5)
        city_var = tk.StringVar()
        city_entry = ttk.Entry(city_frame, textvariable=city_var, width=30)
        city_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Campo opcional para empresa
        company_frame = ttk.Frame(frame)
        company_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(company_frame, text="Empresa:").pack(side=tk.LEFT, padx=5)
        company_var = tk.StringVar()
        company_entry = ttk.Entry(company_frame, textvariable=company_var, width=30)
        company_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Opciones para la categoría
        category_frame = ttk.Frame(frame)
        category_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(category_frame, text="Categoría:").pack(side=tk.LEFT, padx=5)
        
        category_var = tk.StringVar(value="client")
        ttk.Radiobutton(category_frame, text="Cliente", value="client", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(category_frame, text="Comercial", value="commercial", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(category_frame, text="No válido", value="invalid", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        
        # Variable para el motivo (solo para emails no válidos)
        reason_frame = ttk.Frame(frame)
        reason_frame.pack(fill=tk.X, pady=5)
        
        reason_label = ttk.Label(reason_frame, text="Motivo:")
        reason_var = tk.StringVar(value="Añadido manualmente como no válido")
        reason_entry = ttk.Entry(reason_frame, textvariable=reason_var, width=30)
        
        # Función para mostrar/ocultar el campo de motivo según la categoría
        def on_category_change(*args):
            category = category_var.get()
            
            # Mostrar/ocultar campo de motivo solo para emails no válidos
            if category == "invalid":
                reason_label.pack(side=tk.LEFT, padx=5)
                reason_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                
                # Ocultar campos no necesarios para email no válido
                company_frame.pack_forget()
                client_code_frame.pack_forget()
                address_frame.pack_forget()
                postal_code_frame.pack_forget()
                town_frame.pack_forget()
                city_frame.pack_forget()
            else:
                reason_label.pack_forget()
                reason_entry.pack_forget()
                
                # Mostrar campos específicos según la categoría
                if category == "commercial":
                    company_frame.pack(after=email_frame, fill=tk.X, pady=5)
                
                # Mostrar campos comunes
                client_code_frame.pack(after=(company_frame if category == "commercial" else email_frame), fill=tk.X, pady=5)
                address_frame.pack(after=client_code_frame, fill=tk.X, pady=5)
                postal_code_frame.pack(after=address_frame, fill=tk.X, pady=5)
                town_frame.pack(after=postal_code_frame, fill=tk.X, pady=5)
                city_frame.pack(after=town_frame, fill=tk.X, pady=5)
        
        # Asignar función de callback al cambio de categoría
        category_var.trace_add("write", on_category_change)
        
        # Inicializar interfaz según categoría inicial
        on_category_change()
        
        # Función para añadir el contacto
        def on_add():
            # Normalizar valores para prevenir problemas de comparación
            email = (email_var.get() or "").strip().lower()
            name = (name_var.get() or "").strip()
            client_code = (client_code_var.get() or "").strip()
            address = (address_var.get() or "").strip()
            postal_code = (postal_code_var.get() or "").strip()
            town = (town_var.get() or "").strip()
            city = (city_var.get() or "").strip()
            company = (company_var.get() or "").strip()
            reason = (reason_var.get() or "").strip()
            
            # Validar email
            if not email:
                messagebox.showerror("Error", "El campo de email es obligatorio")
                return
            
            if not is_valid_email(email) and category_var.get() != "invalid":
                if messagebox.askyesno("Email no válido", 
                                     "El email ingresado no parece ser válido. ¿Desea continuar de todos modos?"):
                    pass  # Continuar a pesar de ser un email con formato inválido
                else:
                    return
            
            # Verificar si el email ya existe
            existing_category = self.check_existing_email(email)
            
            if existing_category:
                if existing_category == "client":
                    # Si es un cliente existente, intentar actualizar
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    if "updated" in result and result["updated"] and result["changes"]:
                        # Si hay cambios, preguntar si desea aplicarlos
                        changes_message = f"Se han detectado cambios en el cliente {name} <{email}>:\n\n"
                        for change in result["changes"]:
                            changes_message += f"• {change}\n"
                        changes_message += "\n¿Desea aplicar estos cambios?"
                        
                        if messagebox.askyesno("Cambios detectados", changes_message):
                            messagebox.showinfo("Actualización completada", "Los datos del cliente se han actualizado correctamente.")
                            self.update_lists()
                            add_window.destroy()
                        else:
                            # No se aplican los cambios
                            pass
                    else:
                        messagebox.showinfo("Sin cambios", "No se han detectado cambios en los datos del cliente.")
                
                elif existing_category == "commercial":
                    # Si es un contacto comercial existente, intentar actualizar
                    result = self.db_manager.add_commercial_contact_with_changes(
                        name, email, company, client_code, address, postal_code, town, city
                    )
                    if "updated" in result and result["updated"] and result["changes"]:
                        # Si hay cambios, preguntar si desea aplicarlos
                        changes_message = f"Se han detectado cambios en el contacto comercial {name} <{email}>:\n\n"
                        for change in result["changes"]:
                            changes_message += f"• {change}\n"
                        changes_message += "\n¿Desea aplicar estos cambios?"
                        
                        if messagebox.askyesno("Cambios detectados", changes_message):
                            messagebox.showinfo("Actualización completada", "Los datos del contacto comercial se han actualizado correctamente.")
                            self.update_lists()
                            add_window.destroy()
                        else:
                            # No se aplican los cambios
                            pass
                    else:
                        messagebox.showinfo("Sin cambios", "No se han detectado cambios en los datos del contacto comercial.")
                
                elif not self.skip_all_duplicates:
                    # Solo mostrar diálogo para contactos que existen en OTRA categoría (no como cliente)
                    action = self.ask_duplicate_action(email, name, existing_category)
                    
                    if action == "skip":
                        # Omitir este email
                        return
                    elif action == "skip_all":
                        # Omitir todos los duplicados restantes
                        self.skip_all_duplicates = True
                        return
                    elif action == "client":
                        # Reclasificar a cliente
                        self.delete_from_all_categories(email)
                        result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                        if "new" in result and result["new"]:
                            # Agregar a la lista de nuevos contactos
                            self.new_contacts.append(f"{name}_{email}")
                        return
                else:
                    # Si omitir todos está marcado, simplemente saltamos este email
                    return
            else:
                # Si llegamos aquí, es un nuevo contacto (no existe en la base de datos)
                added = False
                if category_var.get() == "client":
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    added = "new" in result and result["new"]
                elif category_var.get() == "commercial":
                    added = self.db_manager.add_commercial_contact_with_changes(
                        name, email, company, client_code, address, postal_code, town, city
                    )
                    added = "new" in added and added["new"]
                elif category_var.get() == "invalid":
                    added = self.db_manager.add_invalid_email(email, name, reason)
                
                if added:
                    # Si es un nuevo cliente o contacto comercial, agregarlo a la lista de nuevos contactos
                    if category_var.get() in ["client", "commercial"]:
                        self.new_contacts.append(f"{name}_{email}")
                    
                    # Actualizar la interfaz
                    self.update_lists()
                    add_window.destroy()
                    messagebox.showinfo("Éxito", "Contacto añadido correctamente")
                else:
                    messagebox.showerror("Error", "No se pudo añadir el contacto")
        
        # Botones
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(buttons_frame, text="Cancelar", 
                 command=add_window.destroy).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(buttons_frame, text="Añadir", 
                 command=on_add, style="success.TButton").pack(side=tk.RIGHT, padx=10)
        
        # Centrar la ventana
        self.center_window(add_window)
        
        # Dar foco al primer campo
        name_entry.focus_set()
    
    def check_existing_email(self, email):
        """Verifica si un email ya existe en alguna tabla."""
        # Comprobar en tabla de clientes
        cursor = self.db_manager.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM clients WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "client"
        
        # Comprobar en tabla de contactos comerciales
        cursor.execute("SELECT COUNT(*) FROM commercial_contacts WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "commercial"
        
        # Comprobar en tabla de emails inválidos
        cursor.execute("SELECT COUNT(*) FROM invalid_emails WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "invalid"
        
        return None  # No existe
    
    def ask_duplicate_action(self, email, name, existing_category):
        """Pregunta al usuario qué hacer con un email duplicado."""
        # Convertir la categoría técnica a un nombre amigable en español
        category_display = {
            "client": "cliente",
            "commercial": "contacto comercial",
            "invalid": "email no válido"
        }.get(existing_category, existing_category)
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Email duplicado")
        dialog.geometry("750x250")  # Aumentado para más espacio horizontal
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
        ttk.Label(frame, text=f"Este email ya existe como {category_display}.").pack(pady=2)
        ttk.Label(frame, text="¿Qué desea hacer?").pack(pady=(10, 5))
        
        # Botones de acción
        actions_frame = ttk.Frame(frame)
        actions_frame.pack(pady=15)
        
        # Función auxiliar para cerrar el diálogo y devolver el resultado
        def set_result_and_close(value):
            result.set(value)
            dialog.destroy()
        
        ttk.Button(actions_frame, text="Omitir", 
                  command=lambda: set_result_and_close("skip"),
                  style="secondary.TButton").pack(side=tk.LEFT, padx=10)
        
        ttk.Button(actions_frame, text="Omitir todos", 
                  command=lambda: set_result_and_close("skip_all"),
                  style="secondary.TButton").pack(side=tk.LEFT, padx=10)
        
        ttk.Button(actions_frame, text="Importar como Cliente", 
                  command=lambda: set_result_and_close("client"),
                  style="success.TButton").pack(side=tk.LEFT, padx=10)
        
        # Centrar la ventana de diálogo
        self.center_window(dialog)
        
        # Esperar a que el usuario responda
        self.root.wait_window(dialog)
        
        # Devolver el resultado
        return result.get() if result.get() else "skip"
    
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
                    # Capturar todos los campos del cliente
                    contacts_to_reclassify.append({
                        "name": values[0],
                        "email": values[1],
                        "client_code": values[3] if len(values) > 3 else "",
                        "address": values[4] if len(values) > 4 else "",
                        "postal_code": values[5] if len(values) > 5 else "",
                        "town": values[6] if len(values) > 6 else "",
                        "city": values[7] if len(values) > 7 else "",
                        "additional_info": values[8] if len(values) > 8 else "",
                    })
                contact_type = "cliente"
        elif current_tab == 1:  # Pestaña de contactos comerciales
            selected_items = self.commercial_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.commercial_tree.item(item, 'values')
                    # Capturar todos los campos del contacto comercial
                    contacts_to_reclassify.append({
                        "name": values[0],
                        "email": values[1],
                        "company": values[2] if len(values) > 2 else "",
                        "client_code": values[4] if len(values) > 4 else "",
                        "address": values[5] if len(values) > 5 else "",
                        "postal_code": values[6] if len(values) > 6 else "",
                        "town": values[7] if len(values) > 7 else "",
                        "city": values[8] if len(values) > 8 else "",
                        "additional_info": values[9] if len(values) > 9 else "",
                    })
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
        target_category_window.geometry("400x220")  # Aumentado para más espacio vertical
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
            # Recuperar todos los campos del contacto
            name = contact.get("name", "")
            email = contact.get("email", "")
            company = contact.get("company", "")
            client_code = contact.get("client_code", "")
            address = contact.get("address", "")
            postal_code = contact.get("postal_code", "")
            town = contact.get("town", "")
            city = contact.get("city", "")
            additional_info = contact.get("additional_info", "")
            
            # Verificar que los valores no sean None
            name = str(name) if name is not None else ""
            email = str(email) if email is not None else ""
            company = str(company) if company is not None else ""
            client_code = str(client_code) if client_code is not None else ""
            address = str(address) if address is not None else ""
            postal_code = str(postal_code) if postal_code is not None else ""
            town = str(town) if town is not None else ""
            city = str(city) if city is not None else ""
            additional_info = str(additional_info) if additional_info is not None else ""
            
            # Depuración
            print(f"DEBUG - Reclasificando contacto: {email}")
            print(f"DEBUG - Tipo original: {current_type}")
            print(f"DEBUG - Nuevo tipo: {new_category_index}")
            print(f"DEBUG - Datos: nombre={name}, email={email}, empresa={company}")
            print(f"DEBUG - Datos adicionales: código={client_code}, dirección={address}, CP={postal_code}")
            print(f"DEBUG - Más datos: población={town}, ciudad={city}, info adicional={additional_info}")
            
            # Crear una conexión a la base de datos si no existe
            if not self.db_manager.conn:
                self.db_manager.create_database()
            
            cursor = self.db_manager.conn.cursor()
            
            # Eliminar de la categoría actual
            try:
                if current_type == "cliente":
                    cursor.execute("DELETE FROM clients WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de clientes: {email}")
                elif current_type == "contacto comercial":
                    cursor.execute("DELETE FROM commercial_contacts WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de contactos comerciales: {email}")
                elif current_type == "email no válido":
                    cursor.execute("DELETE FROM invalid_emails WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de emails no válidos: {email}")
            except Exception as delete_error:
                print(f"ERROR en eliminación: {str(delete_error)}")
            
            # Insertar en la nueva categoría
            if new_category_index == 0:  # Cliente
                try:
                    result = self.db_manager.add_client(
                        name=name, 
                        email=email, 
                        client_code=client_code, 
                        address=address, 
                        postal_code=postal_code, 
                        town=town, 
                        city=city,
                        additional_info=additional_info
                    )
                    print(f"DEBUG - Resultado add_client: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_client: {str(add_error)}")
                    raise
            elif new_category_index == 1:  # Contacto comercial
                try:
                    result = self.db_manager.add_commercial_contact(
                        name=name, 
                        email=email, 
                        company=company, 
                        client_code=client_code, 
                        address=address, 
                        postal_code=postal_code, 
                        town=town, 
                        city=city, 
                        additional_info=additional_info
                    )
                    print(f"DEBUG - Resultado add_commercial_contact: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_commercial_contact: {str(add_error)}")
                    raise
            elif new_category_index == 2:  # Email no válido
                reason = provided_reason or "Marcado manualmente como inválido"
                try:
                    result = self.db_manager.add_invalid_email(email, name, reason)
                    print(f"DEBUG - Resultado add_invalid_email: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_invalid_email: {str(add_error)}")
                    raise
                
            # Confirmar los cambios
            try:
                self.db_manager.conn.commit()
                print(f"DEBUG - Cambios confirmados en la base de datos")
            except Exception as commit_error:
                print(f"ERROR en commit: {str(commit_error)}")
                raise
                
            return True
            
        except Exception as e:
            print(f"ERROR CRÍTICO al reclasificar el contacto {email}: {str(e)}")
            traceback.print_exc()  # Esto imprimirá el traceback completo
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
                                 ["Nombre", "Email", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"])
        elif current_tab == 1:  # Comerciales
            self._export_category("contactos_comerciales", self.db_manager.get_all_commercial_contacts(),
                                 ["Nombre", "Email", "Empresa", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"])
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
            with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
                # Usar punto y coma como delimitador para mejor compatibilidad con Excel español
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(headers)
                
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_data = []
                for row in data:
                    # Si la fila tiene menos columnas que los encabezados, añadir columnas vacías
                    row_list = list(row)
                    while len(row_list) < len(headers):
                        row_list.append("")
                    # Si la fila tiene más columnas que los encabezados, truncar
                    if len(row_list) > len(headers):
                        row_list = row_list[:len(headers)]
                    formatted_data.append(row_list)
                
                writer.writerows(formatted_data)
            
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
            # Definir headers para cada categoría
            client_headers = ["Nombre", "Email", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"]
            commercial_headers = ["Nombre", "Email", "Empresa", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"]
            invalid_headers = ["Email", "Nombre", "Fecha Importación", "Motivo"]
            
            # Exportar clientes
            clients_path = os.path.join(export_dir, "clientes.csv")
            with open(clients_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(client_headers)
                
                clients = self.db_manager.get_all_clients()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_clients = []
                for row in clients:
                    row_list = list(row)
                    while len(row_list) < len(client_headers):
                        row_list.append("")
                    if len(row_list) > len(client_headers):
                        row_list = row_list[:len(client_headers)]
                    formatted_clients.append(row_list)
                
                writer.writerows(formatted_clients)
            
            # Exportar contactos comerciales
            commercial_path = os.path.join(export_dir, "contactos_comerciales.csv")
            with open(commercial_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(commercial_headers)
                
                commercial_contacts = self.db_manager.get_all_commercial_contacts()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_commercial = []
                for row in commercial_contacts:
                    row_list = list(row)
                    while len(row_list) < len(commercial_headers):
                        row_list.append("")
                    if len(row_list) > len(commercial_headers):
                        row_list = row_list[:len(commercial_headers)]
                    formatted_commercial.append(row_list)
                
                writer.writerows(formatted_commercial)
            
            # Exportar emails no válidos
            invalid_path = os.path.join(export_dir, "emails_no_validos.csv")
            with open(invalid_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(invalid_headers)
                
                invalid_emails = self.db_manager.get_all_invalid_emails()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_invalid = []
                for row in invalid_emails:
                    row_list = list(row)
                    while len(row_list) < len(invalid_headers):
                        row_list.append("")
                    if len(row_list) > len(invalid_headers):
                        row_list = row_list[:len(invalid_headers)]
                    formatted_invalid.append(row_list)
                
                writer.writerows(formatted_invalid)
            
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

    def create_column_selector(self, parent, var, label_text, headers):
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(frame, text=label_text, width=20).pack(side=tk.LEFT, padx=5)
        
        # Crear lista de opciones con "-1: No seleccionado" y los índices
        options = [str(i) for i in range(len(headers))]
        if "-1" not in options and label_text != "Email (obligatorio):" and label_text != "Nombre (obligatorio):":
            options.insert(0, "-1")  # -1 para no seleccionado (solo para campos opcionales)
        
        # Crear el combobox
        combo = ttk.Combobox(frame, textvariable=var, values=options, state="readonly", width=10)
        combo.pack(side=tk.LEFT, padx=5)
        
        # Mostrar el encabezado correspondiente al índice seleccionado
        preview_label = ttk.Label(frame, text="")
        preview_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Función para actualizar la etiqueta cuando cambia la selección
        def update_label(*args):
            try:
                idx = int(var.get())
                if idx >= 0 and idx < len(headers):
                    preview_label.config(text=f"-> {headers[idx]}")
                else:
                    preview_label.config(text="-> No seleccionado")
            except (ValueError, IndexError):
                preview_label.config(text="")
        
        # Actualizar la etiqueta inicial
        update_label()
        
        # Vincular la actualización al cambio de variable
        var.trace_add("write", update_label)
        
        return frame

    def reset_search(self):
        """Limpia el campo de búsqueda y actualiza las listas para mostrar todos los contactos."""
        # Limpiar el campo de búsqueda
        self.search_var.set("")
        # Actualizar las listas para mostrar todos los contactos
        self.update_lists()
        # Mensaje opcional para informar al usuario
        messagebox.showinfo("Búsqueda limpiada", "Se han restaurado todas las listas a su estado original.")
    
    def setup_clients_treeview(self):
        """Configura el treeview para clientes regulares."""
        columns = ("name", "email", "date", "client_code", "address", "postal_code", "town", "city", "additional_info")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.clients_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.clients_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                       selectmode="extended")
        
        self.clients_tree.heading("name", text="Nombre")
        self.clients_tree.heading("email", text="Email")
        self.clients_tree.heading("date", text="Fecha Importación")
        self.clients_tree.heading("client_code", text="Código de Cliente")
        self.clients_tree.heading("address", text="Dirección")
        self.clients_tree.heading("postal_code", text="Código Postal")
        self.clients_tree.heading("town", text="Población")
        self.clients_tree.heading("city", text="Ciudad")
        self.clients_tree.heading("additional_info", text="Información Adicional")
        
        self.clients_tree.column("name", width=200)
        self.clients_tree.column("email", width=250)
        self.clients_tree.column("date", width=150)
        self.clients_tree.column("client_code", width=100)
        self.clients_tree.column("address", width=200)
        self.clients_tree.column("postal_code", width=100)
        self.clients_tree.column("town", width=150)
        self.clients_tree.column("city", width=150)
        self.clients_tree.column("additional_info", width=200)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.clients_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.clients_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.clients_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.clients_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar el color verde para los nuevos contactos
        self.clients_tree.tag_configure('new_contact', background='#c8ffb3')
        # Configurar el color verde para las celdas modificadas
        self.clients_tree.tag_configure('modified_cell', background='#a5e88f')
    
    def setup_commercial_treeview(self):
        """Configura el treeview para contactos comerciales."""
        columns = ("name", "email", "company", "date", "client_code", "address", "postal_code", "town", "city", "additional_info")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.commercial_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.commercial_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                          selectmode="extended")
        
        self.commercial_tree.heading("name", text="Nombre")
        self.commercial_tree.heading("email", text="Email")
        self.commercial_tree.heading("company", text="Empresa")
        self.commercial_tree.heading("date", text="Fecha Importación")
        self.commercial_tree.heading("client_code", text="Código de Cliente")
        self.commercial_tree.heading("address", text="Dirección")
        self.commercial_tree.heading("postal_code", text="Código Postal")
        self.commercial_tree.heading("town", text="Población")
        self.commercial_tree.heading("city", text="Ciudad")
        self.commercial_tree.heading("additional_info", text="Información Adicional")
        
        self.commercial_tree.column("name", width=180)
        self.commercial_tree.column("email", width=220)
        self.commercial_tree.column("company", width=180)
        self.commercial_tree.column("date", width=120)
        self.commercial_tree.column("client_code", width=100)
        self.commercial_tree.column("address", width=200)
        self.commercial_tree.column("postal_code", width=100)
        self.commercial_tree.column("town", width=150)
        self.commercial_tree.column("city", width=150)
        self.commercial_tree.column("additional_info", width=200)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.commercial_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.commercial_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.commercial_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.commercial_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
        
        # Configurar el color verde para los nuevos contactos
        self.commercial_tree.tag_configure('new_contact', background='#c8ffb3')
        # Configurar el color verde para las celdas modificadas
        self.commercial_tree.tag_configure('modified_cell', background='#a5e88f')
    
    def setup_invalid_treeview(self):
        """Configura el treeview para emails no válidos."""
        columns = ("email", "name", "date", "reason")
        
        # Crear un frame contenedor para el treeview y las barras de desplazamiento
        treeview_frame = ttk.Frame(self.invalid_frame)
        treeview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Crear el treeview dentro del frame contenedor
        self.invalid_tree = ttk.Treeview(treeview_frame, columns=columns, show="headings", 
                                       selectmode="extended")
        
        self.invalid_tree.heading("email", text="Email")
        self.invalid_tree.heading("name", text="Nombre")
        self.invalid_tree.heading("date", text="Fecha Importación")
        self.invalid_tree.heading("reason", text="Motivo")
        
        self.invalid_tree.column("email", width=250)
        self.invalid_tree.column("name", width=200)
        self.invalid_tree.column("date", width=150)
        self.invalid_tree.column("reason", width=150)
        
        # Barras de desplazamiento
        scrollbar_y = ttk.Scrollbar(treeview_frame, orient=tk.VERTICAL, command=self.invalid_tree.yview)
        scrollbar_x = ttk.Scrollbar(treeview_frame, orient=tk.HORIZONTAL, command=self.invalid_tree.xview)
        
        # Configurar el treeview para usar ambas barras de desplazamiento
        self.invalid_tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)
        
        # Posicionar el treeview y las barras de desplazamiento
        self.invalid_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Hacer que el treeview se expanda con el frame
        treeview_frame.grid_rowconfigure(0, weight=1)
        treeview_frame.grid_columnconfigure(0, weight=1)
    
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
        regular_clients = []
        new_clients = []
        
        # Separar clientes regulares y nuevos
        for client in clients:
            client_id = f"{client[0]}_{client[1]}"  # Usar combinación de nombre y email como ID
            if client_id in self.new_contacts:
                new_clients.append(client)
            else:
                regular_clients.append(client)
        
        # Mostrar clientes regulares
        for client in regular_clients:
            email = client[1]  # El email está en la posición 1
            
            # Comprobar si el cliente está en la lista de contactos modificados
            if email in self.modified_contacts:
                # Si está modificado, insertar con la etiqueta de modificado
                item_id = self.clients_tree.insert("", tk.END, values=client, tags=('modified_cell',))
            else:
                # Si no está modificado, insertar normalmente
                item_id = self.clients_tree.insert("", tk.END, values=client)
        
        # Mostrar clientes nuevos al final con tag especial
        for client in new_clients:
            self.clients_tree.insert("", tk.END, values=client, tags=('new_contact',))
        
        # Actualizar contador de clientes
        total_clients = len(regular_clients) + len(new_clients)
        self.client_count_label.config(text=f"Total Clientes: {total_clients}")
        
        # Obtener y mostrar contactos comerciales
        commercials = self.db_manager.get_all_commercial_contacts()
        regular_commercials = []
        new_commercials = []
        
        # Separar contactos comerciales regulares y nuevos
        for commercial in commercials:
            commercial_id = f"{commercial[0]}_{commercial[1]}"  # Usar combinación de nombre y email como ID
            if commercial_id in self.new_contacts:
                new_commercials.append(commercial)
            else:
                regular_commercials.append(commercial)
        
        # Mostrar contactos comerciales regulares
        for commercial in regular_commercials:
            email = commercial[1]  # El email está en la posición 1
            
            # Comprobar si el contacto comercial está en la lista de contactos modificados
            if email in self.modified_contacts:
                # Si está modificado, insertar con la etiqueta de modificado
                item_id = self.commercial_tree.insert("", tk.END, values=commercial, tags=('modified_cell',))
            else:
                # Si no está modificado, insertar normalmente
                item_id = self.commercial_tree.insert("", tk.END, values=commercial)
        
        # Mostrar contactos comerciales nuevos al final con tag especial
        for commercial in new_commercials:
            self.commercial_tree.insert("", tk.END, values=commercial, tags=('new_contact',))
        
        # Actualizar contador de contactos comerciales
        total_commercials = len(regular_commercials) + len(new_commercials)
        self.commercial_count_label.config(text=f"Total Comerciales: {total_commercials}")
        
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
            
        # Variable para la ventana de progreso
        progress_window = None
        
        # Variable para la ventana de progreso
        progress_window = None
        
        try:
            # Crear ventana de progreso
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Importando CSV")
            progress_window.geometry("400x150")
            progress_window.resizable(False, False)
            self.center_window(progress_window)
            
            # Configurar como ventana modal
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # Añadir mensaje y barra de progreso
            ttk.Label(
                progress_window, 
                text="Importando datos del archivo CSV...\nPor favor espere mientras se procesa el archivo.",
                font=("Segoe UI", 10),
                wraplength=380
            ).pack(pady=15)
            
            progress_bar = ttk.Progressbar(
                progress_window, 
                mode="indeterminate", 
                length=350,
                style="primary.Horizontal.TProgressbar"
            )
            progress_bar.pack(pady=10)
            progress_bar.start(10)
            
            # Actualizar la ventana para asegurar que se muestre
            progress_window.update()
            
            # Reiniciar la opción de omitir todos
            self.skip_all_duplicates = False
            
            # Probar diferentes codificaciones
            encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
            
            for encoding in encodings:
                try:
                    # Intentar abrir el archivo con la codificación actual
                    with open(filepath, 'r', encoding=encoding) as test_file:
                        test_file.readline()  # Leer una línea para verificar
                    
                    # Si llegamos aquí, la codificación funciona
                    break
                except UnicodeDecodeError:
                    if encoding == encodings[-1]:
                        # Si es la última codificación y falla, mostrar error
                        messagebox.showerror("Error", f"No se pudo determinar la codificación del archivo. Prueba guardarlo como UTF-8.")
                        return
                    continue  # Probar con la siguiente codificación
            
            # Primero leer el archivo para determinar las columnas
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')  # Usar ; como delimitador por defecto
                headers = next(csv_reader, None)  # Leer encabezados
            
            if not headers:
                messagebox.showerror("Error", "El archivo CSV no tiene encabezados o está vacío.")
                return
            
            # Crear ventana para seleccionar las columnas
            select_columns_window = tk.Toplevel(self.root)
            select_columns_window.title("Seleccionar columnas")
            # Aumentar el tamaño de la ventana de selección de columnas (era 800x600)
            select_columns_window.geometry("1000x700")
            select_columns_window.resizable(True, True)
            select_columns_window.transient(self.root)
            select_columns_window.grab_set()
            select_columns_window.configure(bg="#f8f4ef")
            
            frame = ttk.Frame(select_columns_window, padding="10")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Información
            ttk.Label(frame, text="Seleccione las columnas que contienen el email y el nombre:", 
                     font=("Helvetica", 10, "bold")).pack(pady=(5, 10))
            
            # Variables para almacenar los índices seleccionados
            # Obtener los valores guardados de la configuración
            saved_email_index = self.db_manager.get_config("email_index", "0")
            saved_name_index = self.db_manager.get_config("name_index", "0")
            saved_client_code_index = self.db_manager.get_config("client_code_index", "-1")
            saved_address_index = self.db_manager.get_config("address_index", "-1")
            saved_postal_code_index = self.db_manager.get_config("postal_code_index", "-1")
            saved_town_index = self.db_manager.get_config("town_index", "-1")
            saved_city_index = self.db_manager.get_config("city_index", "-1")
            
            email_index_var = tk.StringVar(value=saved_email_index)
            name_index_var = tk.StringVar(value=saved_name_index)
            # Additional fields for new columns
            client_code_index_var = tk.StringVar(value=saved_client_code_index)
            address_index_var = tk.StringVar(value=saved_address_index)
            postal_code_index_var = tk.StringVar(value=saved_postal_code_index)
            town_index_var = tk.StringVar(value=saved_town_index)
            city_index_var = tk.StringVar(value=saved_city_index)
            
            # Crear marco para la tabla de vista previa
            preview_frame = ttk.LabelFrame(frame, text="Vista previa del CSV")
            preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            # Crear un canvas y un frame dentro para la tabla
            canvas = tk.Canvas(preview_frame, bg="#f8f4ef")
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=canvas.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            canvas.configure(yscrollcommand=scrollbar.set)
            
            preview_table_frame = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=preview_table_frame, anchor=tk.NW)
            
            # Mostrar las primeras filas del CSV para ayudar en la selección
            num_preview_rows = 5  # Número de filas para vista previa
            
            # Reiniciar el archivo para leer los datos de vista previa
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')
                preview_rows = [next(csv_reader, None)]  # Encabezados
                
                for _ in range(num_preview_rows):
                    try:
                        preview_rows.append(next(csv_reader, None))
                    except StopIteration:
                        break
            
            # Mostrar los datos en una tabla
            max_columns = max(len(row) if row else 0 for row in preview_rows)
            
            # Labels para los índices de columna
            for col in range(max_columns):
                ttk.Label(preview_table_frame, text=f"Col {col}", font=("Helvetica", 9, "bold"), 
                         borderwidth=1, relief="solid", width=15, anchor=tk.CENTER).grid(row=0, column=col, padx=2, pady=2)
            
            # Datos de vista previa
            for row_idx, row in enumerate(preview_rows):
                if row:
                    for col_idx, value in enumerate(row):
                        if col_idx < max_columns:  # Asegurarse de no exceder el número máximo de columnas
                            ttk.Label(preview_table_frame, text=value[:20] + ('...' if len(value) > 20 else ''), 
                                     borderwidth=1, relief="solid", width=15).grid(row=row_idx+1, column=col_idx, padx=2, pady=2)
            
            # Actualizar tamaño del canvas después de agregar widgets
            preview_table_frame.update_idletasks()
            canvas.config(scrollregion=canvas.bbox("all"))
            
            # Selección de columnas
            selection_frame = ttk.Frame(frame)
            selection_frame.pack(fill=tk.X, pady=10)
            
            # Crear selectores para cada campo
            email_selector = self.create_column_selector(selection_frame, email_index_var, "Email (obligatorio):", headers)
            name_selector = self.create_column_selector(selection_frame, name_index_var, "Nombre (obligatorio):", headers)
            
            # Campos adicionales
            client_code_selector = self.create_column_selector(selection_frame, client_code_index_var, "Código Cliente:", headers)
            address_selector = self.create_column_selector(selection_frame, address_index_var, "Dirección:", headers)
            postal_code_selector = self.create_column_selector(selection_frame, postal_code_index_var, "Código Postal:", headers)
            town_selector = self.create_column_selector(selection_frame, town_index_var, "Población:", headers)
            city_selector = self.create_column_selector(selection_frame, city_index_var, "Ciudad:", headers)
            
            # Variable para almacenar la respuesta
            result = {"proceed": False, "email_index": 0, "name_index": 0, "client_code_index": -1, "address_index": -1, "postal_code_index": -1, "town_index": -1, "city_index": -1}
            
            # Función para procesar la selección
            def on_submit():
                result["proceed"] = True
                result["email_index"] = email_index_var.get()
                result["name_index"] = name_index_var.get()
                result["client_code_index"] = client_code_index_var.get()
                result["address_index"] = address_index_var.get()
                result["postal_code_index"] = postal_code_index_var.get()
                result["town_index"] = town_index_var.get()
                result["city_index"] = city_index_var.get()
                select_columns_window.destroy()
            
            # Botones
            buttons_frame = ttk.Frame(frame)
            buttons_frame.pack(fill=tk.X, pady=10)
            
            ttk.Button(buttons_frame, text="Cancelar", 
                      command=select_columns_window.destroy).pack(side=tk.LEFT, padx=10)
            
            ttk.Button(buttons_frame, text="Importar", 
                      command=on_submit, style="success.TButton").pack(side=tk.RIGHT, padx=10)
            
            # Centrar la ventana
            self.center_window(select_columns_window)
            
            # Esperar a que se cierre la ventana
            self.root.wait_window(select_columns_window)
            
            if not result["proceed"]:
                return
            
            # Guardar la configuración para futuras importaciones
            self.db_manager.save_config("email_index", str(result["email_index"]))
            self.db_manager.save_config("name_index", str(result["name_index"]))
            self.db_manager.save_config("client_code_index", str(result["client_code_index"]))
            self.db_manager.save_config("address_index", str(result["address_index"]))
            self.db_manager.save_config("postal_code_index", str(result["postal_code_index"]))
            self.db_manager.save_config("town_index", str(result["town_index"]))
            self.db_manager.save_config("city_index", str(result["city_index"]))
            
            # Asegurarnos de que la conexión a la base de datos está activa
            if self.db_manager.conn is None or not hasattr(self.db_manager, 'conn'):
                self.db_manager.create_database()
            
            # Limpiar la lista de nuevos contactos antes de importar
            self.new_contacts = []
            
            # Convertir índices a enteros
            email_index = int(result["email_index"])
            name_index = int(result["name_index"])
            client_code_index = int(result["client_code_index"])
            address_index = int(result["address_index"])
            postal_code_index = int(result["postal_code_index"])
            town_index = int(result["town_index"])
            city_index = int(result["city_index"])
            
            # SOLUCIÓN: Leer todo el CSV y luego procesar por email para garantizar coherencia
            # Primero leemos todo el CSV y lo organizamos por email
            csv_data = {}
            
            with open(filepath, 'r', encoding=encoding) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')
                # Eliminar la siguiente línea para no saltar encabezados
                # next(csv_reader, None)  # Saltar encabezados
                
                # Contadores para depuración
                skipped_insufficient_columns = 0
                skipped_empty_email = 0
                
                for row in csv_reader:
                    # Verificar que la fila tenga suficientes columnas
                    max_required_idx = max(email_index, name_index, 
                                       client_code_index if client_code_index > 0 else 0,
                                       address_index if address_index > 0 else 0,
                                       postal_code_index if postal_code_index > 0 else 0,
                                       town_index if town_index > 0 else 0,
                                       city_index if city_index > 0 else 0)
                    
                    if len(row) <= max_required_idx:
                        skipped_insufficient_columns += 1
                        if self.db_manager.debug_mode:
                            print(f"DEBUG: Saltando fila con columnas insuficientes. Tiene {len(row)} columnas, se necesitan {max_required_idx+1}.")
                            if len(row) > 0:
                                print(f"Contenido de la fila: {row}")
                        continue  # Saltar filas que no tienen suficientes columnas
                    
                    email = row[email_index].strip().lower() if email_index < len(row) else ""
                    
                    if not email:
                        skipped_empty_email += 1
                        if self.db_manager.debug_mode:
                            print(f"DEBUG: Saltando fila sin email. Contenido: {row}")
                        continue  # Saltar filas sin email
                    
                    # Guardar los datos de esta fila organizados por email
                    csv_data[email] = {
                        'name': row[name_index].strip() if name_index < len(row) else "",
                        'client_code': row[client_code_index].strip() if client_code_index >= 0 and client_code_index < len(row) else "",
                        'address': row[address_index].strip() if address_index >= 0 and address_index < len(row) else "",
                        'postal_code': row[postal_code_index].strip() if postal_code_index >= 0 and postal_code_index < len(row) else "",
                        'town': row[town_index].strip() if town_index >= 0 and town_index < len(row) else "",
                        'city': row[city_index].strip() if city_index >= 0 and city_index < len(row) else ""
                    }
            
            # Contadores para el informe final
            new_clients = 0
            updated_clients = 0
            unchanged_clients = 0
            duplicate_contacts = 0
            clients_with_changes = []
            
            # Ahora procesamos cada email asegurándonos que la asociación es correcta
            for email, data in csv_data.items():
                name = data['name']
                client_code = data['client_code']
                address = data['address']
                postal_code = data['postal_code']
                town = data['town']
                city = data['city']
                
                # Verificar si el email ya existe en alguna tabla
                existing_category = self.check_existing_email(email)
                
                if existing_category:
                    # Comprobar si existe en la base de datos y verificar si hay cambios
                    if existing_category == "client":
                        # Si es un cliente existente, intentar actualizar o detectar cambios
                        result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                        
                        # Procesar el resultado
                        if "new" in result and result["new"]:
                            new_clients += 1
                            self.new_contacts.append(f"{name}_{email}")
                        
                        elif "updated" in result:
                            if result["updated"] and result["changes"]:
                                # Solo mostrar diálogo si hay cambios reales
                                # Preguntar al usuario si desea aplicar los cambios
                                changes_message = f"Se han detectado cambios en el cliente {name} <{email}>:\n\n"
                                for change in result["changes"]:
                                    changes_message += f"• {change}\n"
                                changes_message += "\n¿Desea aplicar estos cambios?"
                                
                                # Añadir opción para activar depuración si hay muchos cambios
                                if len(result["changes"]) > 3:
                                    changes_message += "\n\nNota: Si cree que estos cambios no son correctos, puede activar el modo de depuración en la parte inferior derecha de la ventana para obtener más información."
                                
                                if messagebox.askyesno("Cambios detectados", changes_message):
                                    # Si el usuario acepta, actualizar el contador
                                    updated_clients += 1
                                    # Guardar los detalles de los cambios
                                    clients_with_changes.append({
                                        "name": name,
                                        "email": email,
                                        "changes": result["changes"]
                                    })
                                    
                                    # Guardar los campos modificados para colorearlos después
                                    modified_fields = []
                                    for change in result["changes"]:
                                        parts = change.split(": ")
                                        if len(parts) >= 2:
                                            field_name = parts[0]
                                            # Mapear los nombres de campo en español a los nombres de columna en inglés
                                            field_mapping = {
                                                "nombre": "name",
                                                "código cliente": "client_code",
                                                "dirección": "address",
                                                "código postal": "postal_code",
                                                "población": "town",
                                                "ciudad": "city"
                                            }
                                            if field_name in field_mapping:
                                                modified_fields.append(field_mapping[field_name])
                                    
                                    # Guardar en el diccionario de contactos modificados
                                    self.modified_contacts[email] = modified_fields
                                else:
                                    # Si el usuario rechaza, revertir los cambios en la base de datos
                                    cursor = self.db_manager.conn.cursor()
                                    cursor.execute("SELECT id FROM clients WHERE email = ?", (email,))
                                    client_id = cursor.fetchone()[0]
                                    
                                    # Buscar los valores originales
                                    for change in result["changes"]:
                                        parts = change.split(": ")
                                        if len(parts) >= 2:
                                            field = parts[0]
                                            values_part = parts[1]
                                            value_parts = values_part.split(" -> ")
                                            if len(value_parts) >= 1:
                                                original = value_parts[0].strip("'")
                                                
                                                # Actualizar los campos que han cambiado
                                                if "nombre" in field:
                                                    cursor.execute("UPDATE clients SET name = ? WHERE id = ?", (original, client_id))
                                                elif "código cliente" in field:
                                                    cursor.execute("UPDATE clients SET client_code = ? WHERE id = ?", (original, client_id))
                                                elif "dirección" in field:
                                                    cursor.execute("UPDATE clients SET address = ? WHERE id = ?", (original, client_id))
                                                elif "código postal" in field:
                                                    cursor.execute("UPDATE clients SET postal_code = ? WHERE id = ?", (original, client_id))
                                                elif "población" in field:
                                                    cursor.execute("UPDATE clients SET town = ? WHERE id = ?", (original, client_id))
                                                elif "ciudad" in field:
                                                    cursor.execute("UPDATE clients SET city = ? WHERE id = ?", (original, client_id))
                                    
                                    self.db_manager.conn.commit()
                                    unchanged_clients += 1
                            else:
                                # Si no hay cambios, incrementar el contador de sin cambios
                                unchanged_clients += 1
                        elif "error" in result:
                            print(f"Error al procesar cliente {name}, {email}: {result['error']}")
                    elif existing_category == "commercial":
                        # Si es un contacto comercial existente, utilizar el nuevo método con detección de cambios
                        result = self.db_manager.add_commercial_contact_with_changes(
                            name, email, "", client_code, address, postal_code, town, city
                        )
                        
                        if "new" in result and result["new"]:
                            # No debería ser un nuevo contacto si ya existe
                            pass
                        elif "updated" in result:
                            if result["updated"] and result["changes"]:
                                # Filtrar solo los cambios reales en los datos
                                real_changes = []
                                for change in result["changes"]:
                                    # Ignorar cambios de categoría
                                    if not any(cat in change.lower() for cat in ["categoría", "category"]):
                                        real_changes.append(change)
                                
                                # Verificar si el contacto acaba de ser reclasificado recientemente
                                cursor = self.db_manager.conn.cursor()
                                
                                # Solo considerar como "recién movido" si fue añadido hace menos de 10 minutos
                                # Lo que generalmente indica que fue parte de la misma operación
                                cursor.execute("SELECT imported_date FROM commercial_contacts WHERE email = ?", (email,))
                                commercial_date = cursor.fetchone()
                                contact_just_moved = False
                                
                                if commercial_date:
                                    last_date = commercial_date[0]
                                    now = datetime.now()
                                    
                                    if isinstance(last_date, str):
                                        try:
                                            last_date = datetime.strptime(last_date, "%Y-%m-%d %H:%M:%S")
                                            # Si fue agregado en los últimos 10 minutos, probablemente es 
                                            # parte de la misma operación de reclasificación
                                            if (now - last_date).total_seconds() < 600:  # 10 minutos
                                                # Buscar si existía antes en otra tabla (cliente)
                                                cursor.execute("SELECT COUNT(*) FROM clients WHERE email = ?", (email,))
                                                client_count = cursor.fetchone()[0]
                                                if client_count > 0:
                                                    # Si existe en ambas tablas, es una reclasificación en progreso
                                                    contact_just_moved = True
                                        except ValueError:
                                            pass
                                
                                # Si hay cambios importantes en el contacto (no solo cambio de categoría)
                                if real_changes:
                                    # Preguntar al usuario si desea aplicar los cambios
                                    changes_message = f"Se han detectado cambios en el contacto comercial {name} <{email}>:\n\n"
                                    for change in real_changes:
                                        changes_message += f"• {change}\n"
                                    changes_message += "\n¿Desea aplicar estos cambios?"
                                    
                                    # Añadir opción para activar depuración si hay muchos cambios
                                    if len(real_changes) > 3:
                                        changes_message += "\n\nNota: Si cree que estos cambios no son correctos, puede activar el modo de depuración en la parte inferior derecha de la ventana para obtener más información."
                                    
                                    if messagebox.askyesno("Cambios detectados", changes_message):
                                        # Si el usuario acepta, actualizar el contador
                                        updated_clients += 1
                                        # Guardar los detalles de los cambios
                                        clients_with_changes.append({
                                            "name": name,
                                            "email": email,
                                            "changes": real_changes
                                        })
                                        
                                        # Guardar los campos modificados para colorearlos después
                                        modified_fields = []
                                        for change in real_changes:
                                            parts = change.split(": ")
                                            if len(parts) >= 2:
                                                field_name = parts[0]
                                                # Mapear los nombres de campo en español a los nombres de columna en inglés
                                                field_mapping = {
                                                    "nombre": "name",
                                                    "empresa": "company",
                                                    "código cliente": "client_code",
                                                    "dirección": "address",
                                                    "código postal": "postal_code",
                                                    "población": "town",
                                                    "ciudad": "city",
                                                    "información adicional": "additional_info"
                                                }
                                                if field_name in field_mapping:
                                                    modified_fields.append(field_mapping[field_name])
                                        
                                        # Guardar en el diccionario de contactos modificados
                                        self.modified_contacts[email] = modified_fields
                                else:
                                    # No hay cambios reales en los datos
                                    unchanged_clients += 1
                            else:
                                # No hay cambios según la BD
                                unchanged_clients += 1
                        elif "error" in result:
                            print(f"Error al procesar contacto comercial {name}, {email}: {result['error']}")
                    elif existing_category == "invalid":
                        # Si es un email no válido, simplemente omitirlo
                        duplicate_contacts += 1
                        continue
                    elif not self.skip_all_duplicates:
                        # Solo mostrar diálogo para contactos que existen en OTRA categoría (no como cliente)
                        action = self.ask_duplicate_action(email, name, existing_category)
                        
                        if action == "skip":
                            # Omitir este email
                            continue
                        elif action == "skip_all":
                            # Omitir todos los duplicados restantes
                            self.skip_all_duplicates = True
                            continue
                        elif action == "client":
                            # Reclasificar a cliente
                            self.delete_from_all_categories(email)
                            result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                            if "new" in result and result["new"]:
                                # Agregar a la lista de nuevos contactos
                                self.new_contacts.append(f"{name}_{email}")
                            continue
                    else:
                        # Si omitir todos está marcado, simplemente saltamos este email
                        continue
                else:
                    # Si llegamos aquí, es un nuevo contacto (no existe en la base de datos)
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    if "new" in result and result["new"]:
                        new_clients += 1
                        # Agregar a la lista de nuevos contactos
                        self.new_contacts.append(f"{name}_{email}")
            
            self.update_lists()
            
            # Preparar mensaje detallado
            detailed_message = f"Se han añadido {new_clients} nuevos clientes.\n"\
                              f"Se han actualizado {updated_clients} clientes existentes.\n"\
                              f"Se han omitido {unchanged_clients} clientes sin cambios.\n"\
                              f"Se han omitido {duplicate_contacts} contactos duplicados."
            
            # Añadir información de depuración sobre filas saltadas
            if self.db_manager.debug_mode:
                total_rows = skipped_insufficient_columns + skipped_empty_email + len(csv_data)
                detailed_message += f"\n\nInformación de depuración:\n"\
                                   f"- Total de filas en CSV: {total_rows}\n"\
                                   f"- Filas procesadas correctamente: {len(csv_data)}\n"\
                                   f"- Filas saltadas por columnas insuficientes: {skipped_insufficient_columns}\n"\
                                   f"- Filas saltadas por email vacío: {skipped_empty_email}\n"\
                                   f"- Duplicados (mismo email): {total_rows - skipped_insufficient_columns - skipped_empty_email - len(csv_data)}"
            
            # Si hay clientes actualizados, mostrar los detalles de los cambios
            if updated_clients > 0 and clients_with_changes:
                changes_details = "\n\nDetalles de los cambios:\n"
                for client in clients_with_changes:
                    changes_details += f"\n- {client['name']} ({client['email']}):\n"
                    for change in client['changes']:
                        changes_details += f"  • {change}\n"
                
                # Si el mensaje es muy largo, preguntar si quiere ver los detalles
                if len(changes_details) > 500:
                    show_details = messagebox.askyesno(
                        "Cambios Detallados",
                        f"{detailed_message}\n\n¿Desea ver los detalles de todos los cambios?"
                    )
                    if show_details:
                        # Mostrar detalles en una ventana de texto desplazable
                        details_window = tk.Toplevel(self.root)
                        details_window.title("Detalles de Cambios")
                        details_window.geometry("700x500")
                        details_window.transient(self.root)
                        details_window.grab_set()
                        
                        # Frame principal
                        frame = ttk.Frame(details_window, padding=10)
                        frame.pack(fill=tk.BOTH, expand=True)
                        
                        # Título
                        ttk.Label(
                            frame, 
                            text="Detalles de contactos actualizados", 
                            font=("Helvetica", 12, "bold")
                        ).pack(anchor=tk.W, pady=(0, 10))
                        
                        # Área de texto con scroll
                        text_frame = ttk.Frame(frame)
                        text_frame.pack(fill=tk.BOTH, expand=True)
                        
                        text = tk.Text(text_frame, wrap=tk.WORD)
                        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                        
                        scrollbar = ttk.Scrollbar(text_frame, command=text.yview)
                        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                        
                        text.configure(yscrollcommand=scrollbar.set)
                        text.insert(tk.END, changes_details)
                        text.configure(state="disabled")  # Solo lectura
                        
                        # Botón para cerrar
                        ttk.Button(
                            frame, 
                            text="Cerrar", 
                            command=details_window.destroy
                        ).pack(pady=10)
                        
                        # Centrar la ventana relativa a la ventana principal
                        self.center_window(details_window)
                else:
                    # Si el mensaje no es muy largo, mostrar todo junto
                    messagebox.showinfo(
                        "Importación Completada", 
                        f"{detailed_message}\n{changes_details}"
                    )
            else:
                # Si no hay cambios, mostrar solo el mensaje básico
                messagebox.showinfo("Importación Completada", detailed_message)
            
            # Limpiar variables que controlan la importación para próximos usos
            self.skip_all_duplicates = False
            
            # No limpiamos self.new_contacts para que se muestren resaltados hasta la próxima importación
            # En cambio, limpiamos los contactos modificados ya que esos sí tienen el efecto permanente en la BD
            # y ya han sido vistos en el informe que se acaba de mostrar
            self.modified_contacts = {}
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar el archivo: {str(e)}")
            
        finally:
            # Cerrar la ventana de progreso SIEMPRE, incluso si hay excepciones
            if 'progress_window' in locals() and progress_window and progress_window.winfo_exists():
                progress_window.destroy()
    
    def add_contact_manually(self):
        """Permite agregar un contacto manualmente."""
        # Crear ventana para añadir contacto
        add_window = tk.Toplevel(self.root)
        add_window.title("Agregar contacto manualmente")
        add_window.geometry("500x550")
        add_window.resizable(False, False)
        add_window.transient(self.root)
        add_window.grab_set()
        add_window.configure(bg="#f8f4ef")
        
        frame = ttk.Frame(add_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Ingrese los datos del nuevo contacto:", 
                 font=("Helvetica", 10, "bold")).pack(pady=(5, 15))
        
        # Campos para nombre y email
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(name_frame, text="Nombre:").pack(side=tk.LEFT, padx=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        email_frame = ttk.Frame(frame)
        email_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(email_frame, text="Email:").pack(side=tk.LEFT, padx=5)
        email_var = tk.StringVar()
        email_entry = ttk.Entry(email_frame, textvariable=email_var, width=30)
        email_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Nuevos campos
        # Código de cliente
        client_code_frame = ttk.Frame(frame)
        client_code_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(client_code_frame, text="Código de Cliente:").pack(side=tk.LEFT, padx=5)
        client_code_var = tk.StringVar()
        client_code_entry = ttk.Entry(client_code_frame, textvariable=client_code_var, width=30)
        client_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Dirección
        address_frame = ttk.Frame(frame)
        address_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(address_frame, text="Dirección:").pack(side=tk.LEFT, padx=5)
        address_var = tk.StringVar()
        address_entry = ttk.Entry(address_frame, textvariable=address_var, width=30)
        address_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Código postal
        postal_code_frame = ttk.Frame(frame)
        postal_code_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(postal_code_frame, text="Código Postal:").pack(side=tk.LEFT, padx=5)
        postal_code_var = tk.StringVar()
        postal_code_entry = ttk.Entry(postal_code_frame, textvariable=postal_code_var, width=30)
        postal_code_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Población
        town_frame = ttk.Frame(frame)
        town_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(town_frame, text="Población:").pack(side=tk.LEFT, padx=5)
        town_var = tk.StringVar()
        town_entry = ttk.Entry(town_frame, textvariable=town_var, width=30)
        town_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Ciudad
        city_frame = ttk.Frame(frame)
        city_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(city_frame, text="Ciudad:").pack(side=tk.LEFT, padx=5)
        city_var = tk.StringVar()
        city_entry = ttk.Entry(city_frame, textvariable=city_var, width=30)
        city_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Campo opcional para empresa
        company_frame = ttk.Frame(frame)
        company_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(company_frame, text="Empresa:").pack(side=tk.LEFT, padx=5)
        company_var = tk.StringVar()
        company_entry = ttk.Entry(company_frame, textvariable=company_var, width=30)
        company_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Opciones para la categoría
        category_frame = ttk.Frame(frame)
        category_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(category_frame, text="Categoría:").pack(side=tk.LEFT, padx=5)
        
        category_var = tk.StringVar(value="client")
        ttk.Radiobutton(category_frame, text="Cliente", value="client", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(category_frame, text="Comercial", value="commercial", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(category_frame, text="No válido", value="invalid", 
                       variable=category_var).pack(side=tk.LEFT, padx=10)
        
        # Variable para el motivo (solo para emails no válidos)
        reason_frame = ttk.Frame(frame)
        reason_frame.pack(fill=tk.X, pady=5)
        
        reason_label = ttk.Label(reason_frame, text="Motivo:")
        reason_var = tk.StringVar(value="Añadido manualmente como no válido")
        reason_entry = ttk.Entry(reason_frame, textvariable=reason_var, width=30)
        
        # Función para mostrar/ocultar el campo de motivo según la categoría
        def on_category_change(*args):
            category = category_var.get()
            
            # Mostrar/ocultar campo de motivo solo para emails no válidos
            if category == "invalid":
                reason_label.pack(side=tk.LEFT, padx=5)
                reason_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
                
                # Ocultar campos no necesarios para email no válido
                company_frame.pack_forget()
                client_code_frame.pack_forget()
                address_frame.pack_forget()
                postal_code_frame.pack_forget()
                town_frame.pack_forget()
                city_frame.pack_forget()
            else:
                reason_label.pack_forget()
                reason_entry.pack_forget()
                
                # Mostrar campos específicos según la categoría
                if category == "commercial":
                    company_frame.pack(after=email_frame, fill=tk.X, pady=5)
                
                # Mostrar campos comunes
                client_code_frame.pack(after=(company_frame if category == "commercial" else email_frame), fill=tk.X, pady=5)
                address_frame.pack(after=client_code_frame, fill=tk.X, pady=5)
                postal_code_frame.pack(after=address_frame, fill=tk.X, pady=5)
                town_frame.pack(after=postal_code_frame, fill=tk.X, pady=5)
                city_frame.pack(after=town_frame, fill=tk.X, pady=5)
        
        # Asignar función de callback al cambio de categoría
        category_var.trace_add("write", on_category_change)
        
        # Inicializar interfaz según categoría inicial
        on_category_change()
        
        # Función para añadir el contacto
        def on_add():
            # Normalizar valores para prevenir problemas de comparación
            email = (email_var.get() or "").strip().lower()
            name = (name_var.get() or "").strip()
            client_code = (client_code_var.get() or "").strip()
            address = (address_var.get() or "").strip()
            postal_code = (postal_code_var.get() or "").strip()
            town = (town_var.get() or "").strip()
            city = (city_var.get() or "").strip()
            company = (company_var.get() or "").strip()
            reason = (reason_var.get() or "").strip()
            
            # Validar email
            if not email:
                messagebox.showerror("Error", "El campo de email es obligatorio")
                return
            
            if not is_valid_email(email) and category_var.get() != "invalid":
                if messagebox.askyesno("Email no válido", 
                                     "El email ingresado no parece ser válido. ¿Desea continuar de todos modos?"):
                    pass  # Continuar a pesar de ser un email con formato inválido
                else:
                    return
            
            # Verificar si el email ya existe
            existing_category = self.check_existing_email(email)
            
            if existing_category:
                if existing_category == "client":
                    # Si es un cliente existente, intentar actualizar
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    if "updated" in result and result["updated"] and result["changes"]:
                        # Si hay cambios, preguntar si desea aplicarlos
                        changes_message = f"Se han detectado cambios en el cliente {name} <{email}>:\n\n"
                        for change in result["changes"]:
                            changes_message += f"• {change}\n"
                        changes_message += "\n¿Desea aplicar estos cambios?"
                        
                        if messagebox.askyesno("Cambios detectados", changes_message):
                            messagebox.showinfo("Actualización completada", "Los datos del cliente se han actualizado correctamente.")
                            self.update_lists()
                            add_window.destroy()
                        else:
                            # No se aplican los cambios
                            pass
                    else:
                        messagebox.showinfo("Sin cambios", "No se han detectado cambios en los datos del cliente.")
                
                elif existing_category == "commercial":
                    # Si es un contacto comercial existente, intentar actualizar
                    result = self.db_manager.add_commercial_contact_with_changes(
                        name, email, company, client_code, address, postal_code, town, city
                    )
                    if "updated" in result and result["updated"] and result["changes"]:
                        # Si hay cambios, preguntar si desea aplicarlos
                        changes_message = f"Se han detectado cambios en el contacto comercial {name} <{email}>:\n\n"
                        for change in result["changes"]:
                            changes_message += f"• {change}\n"
                        changes_message += "\n¿Desea aplicar estos cambios?"
                        
                        if messagebox.askyesno("Cambios detectados", changes_message):
                            messagebox.showinfo("Actualización completada", "Los datos del contacto comercial se han actualizado correctamente.")
                            self.update_lists()
                            add_window.destroy()
                        else:
                            # No se aplican los cambios
                            pass
                    else:
                        messagebox.showinfo("Sin cambios", "No se han detectado cambios en los datos del contacto comercial.")
                
                elif not self.skip_all_duplicates:
                    # Solo mostrar diálogo para contactos que existen en OTRA categoría (no como cliente)
                    action = self.ask_duplicate_action(email, name, existing_category)
                    
                    if action == "skip":
                        # Omitir este email
                        return
                    elif action == "skip_all":
                        # Omitir todos los duplicados restantes
                        self.skip_all_duplicates = True
                        return
                    elif action == "client":
                        # Reclasificar a cliente
                        self.delete_from_all_categories(email)
                        result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                        if "new" in result and result["new"]:
                            # Agregar a la lista de nuevos contactos
                            self.new_contacts.append(f"{name}_{email}")
                        return
                else:
                    # Si omitir todos está marcado, simplemente saltamos este email
                    return
            else:
                # Si llegamos aquí, es un nuevo contacto (no existe en la base de datos)
                added = False
                if category_var.get() == "client":
                    result = self.db_manager.add_client(name, email, client_code, address, postal_code, town, city)
                    added = "new" in result and result["new"]
                elif category_var.get() == "commercial":
                    added = self.db_manager.add_commercial_contact_with_changes(
                        name, email, company, client_code, address, postal_code, town, city
                    )
                    added = "new" in added and added["new"]
                elif category_var.get() == "invalid":
                    added = self.db_manager.add_invalid_email(email, name, reason)
                
                if added:
                    # Si es un nuevo cliente o contacto comercial, agregarlo a la lista de nuevos contactos
                    if category_var.get() in ["client", "commercial"]:
                        self.new_contacts.append(f"{name}_{email}")
                    
                    # Actualizar la interfaz
                    self.update_lists()
                    add_window.destroy()
                    messagebox.showinfo("Éxito", "Contacto añadido correctamente")
                else:
                    messagebox.showerror("Error", "No se pudo añadir el contacto")
        
        # Botones
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, pady=15)
        
        ttk.Button(buttons_frame, text="Cancelar", 
                 command=add_window.destroy).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(buttons_frame, text="Añadir", 
                 command=on_add, style="success.TButton").pack(side=tk.RIGHT, padx=10)
        
        # Centrar la ventana
        self.center_window(add_window)
        
        # Dar foco al primer campo
        name_entry.focus_set()
    
    def check_existing_email(self, email):
        """Verifica si un email ya existe en alguna tabla."""
        # Comprobar en tabla de clientes
        cursor = self.db_manager.conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM clients WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "client"
        
        # Comprobar en tabla de contactos comerciales
        cursor.execute("SELECT COUNT(*) FROM commercial_contacts WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "commercial"
        
        # Comprobar en tabla de emails inválidos
        cursor.execute("SELECT COUNT(*) FROM invalid_emails WHERE email = ?", (email,))
        count = cursor.fetchone()[0]
        if count > 0:
            return "invalid"
        
        return None  # No existe
    
    def ask_duplicate_action(self, email, name, existing_category):
        """Pregunta al usuario qué hacer con un email duplicado."""
        # Convertir la categoría técnica a un nombre amigable en español
        category_display = {
            "client": "cliente",
            "commercial": "contacto comercial",
            "invalid": "email no válido"
        }.get(existing_category, existing_category)
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Email duplicado")
        dialog.geometry("750x250")  # Aumentado para más espacio horizontal
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
        ttk.Label(frame, text=f"Este email ya existe como {category_display}.").pack(pady=2)
        ttk.Label(frame, text="¿Qué desea hacer?").pack(pady=(10, 5))
        
        # Botones de acción
        actions_frame = ttk.Frame(frame)
        actions_frame.pack(pady=15)
        
        # Función auxiliar para cerrar el diálogo y devolver el resultado
        def set_result_and_close(value):
            result.set(value)
            dialog.destroy()
        
        ttk.Button(actions_frame, text="Omitir", 
                  command=lambda: set_result_and_close("skip"),
                  style="secondary.TButton").pack(side=tk.LEFT, padx=10)
        
        ttk.Button(actions_frame, text="Omitir todos", 
                  command=lambda: set_result_and_close("skip_all"),
                  style="secondary.TButton").pack(side=tk.LEFT, padx=10)
        
        ttk.Button(actions_frame, text="Importar como Cliente", 
                  command=lambda: set_result_and_close("client"),
                  style="success.TButton").pack(side=tk.LEFT, padx=10)
        
        # Centrar la ventana de diálogo
        self.center_window(dialog)
        
        # Esperar a que el usuario responda
        self.root.wait_window(dialog)
        
        # Devolver el resultado
        return result.get() if result.get() else "skip"
    
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
                    # Capturar todos los campos del cliente
                    contacts_to_reclassify.append({
                        "name": values[0],
                        "email": values[1],
                        "client_code": values[3] if len(values) > 3 else "",
                        "address": values[4] if len(values) > 4 else "",
                        "postal_code": values[5] if len(values) > 5 else "",
                        "town": values[6] if len(values) > 6 else "",
                        "city": values[7] if len(values) > 7 else "",
                        "additional_info": values[8] if len(values) > 8 else "",
                    })
                contact_type = "cliente"
        elif current_tab == 1:  # Pestaña de contactos comerciales
            selected_items = self.commercial_tree.selection()
            if selected_items:
                for item in selected_items:
                    values = self.commercial_tree.item(item, 'values')
                    # Capturar todos los campos del contacto comercial
                    contacts_to_reclassify.append({
                        "name": values[0],
                        "email": values[1],
                        "company": values[2] if len(values) > 2 else "",
                        "client_code": values[4] if len(values) > 4 else "",
                        "address": values[5] if len(values) > 5 else "",
                        "postal_code": values[6] if len(values) > 6 else "",
                        "town": values[7] if len(values) > 7 else "",
                        "city": values[8] if len(values) > 8 else "",
                        "additional_info": values[9] if len(values) > 9 else "",
                    })
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
        target_category_window.geometry("400x220")  # Aumentado para más espacio vertical
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
            # Recuperar todos los campos del contacto
            name = contact.get("name", "")
            email = contact.get("email", "")
            company = contact.get("company", "")
            client_code = contact.get("client_code", "")
            address = contact.get("address", "")
            postal_code = contact.get("postal_code", "")
            town = contact.get("town", "")
            city = contact.get("city", "")
            additional_info = contact.get("additional_info", "")
            
            # Verificar que los valores no sean None
            name = str(name) if name is not None else ""
            email = str(email) if email is not None else ""
            company = str(company) if company is not None else ""
            client_code = str(client_code) if client_code is not None else ""
            address = str(address) if address is not None else ""
            postal_code = str(postal_code) if postal_code is not None else ""
            town = str(town) if town is not None else ""
            city = str(city) if city is not None else ""
            additional_info = str(additional_info) if additional_info is not None else ""
            
            # Depuración
            print(f"DEBUG - Reclasificando contacto: {email}")
            print(f"DEBUG - Tipo original: {current_type}")
            print(f"DEBUG - Nuevo tipo: {new_category_index}")
            print(f"DEBUG - Datos: nombre={name}, email={email}, empresa={company}")
            print(f"DEBUG - Datos adicionales: código={client_code}, dirección={address}, CP={postal_code}")
            print(f"DEBUG - Más datos: población={town}, ciudad={city}, info adicional={additional_info}")
            
            # Crear una conexión a la base de datos si no existe
            if not self.db_manager.conn:
                self.db_manager.create_database()
            
            cursor = self.db_manager.conn.cursor()
            
            # Eliminar de la categoría actual
            try:
                if current_type == "cliente":
                    cursor.execute("DELETE FROM clients WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de clientes: {email}")
                elif current_type == "contacto comercial":
                    cursor.execute("DELETE FROM commercial_contacts WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de contactos comerciales: {email}")
                elif current_type == "email no válido":
                    cursor.execute("DELETE FROM invalid_emails WHERE email=?", (email,))
                    print(f"DEBUG - Eliminado de emails no válidos: {email}")
            except Exception as delete_error:
                print(f"ERROR en eliminación: {str(delete_error)}")
            
            # Insertar en la nueva categoría
            if new_category_index == 0:  # Cliente
                try:
                    result = self.db_manager.add_client(
                        name=name, 
                        email=email, 
                        client_code=client_code, 
                        address=address, 
                        postal_code=postal_code, 
                        town=town, 
                        city=city,
                        additional_info=additional_info
                    )
                    print(f"DEBUG - Resultado add_client: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_client: {str(add_error)}")
                    raise
            elif new_category_index == 1:  # Contacto comercial
                try:
                    result = self.db_manager.add_commercial_contact(
                        name=name, 
                        email=email, 
                        company=company, 
                        client_code=client_code, 
                        address=address, 
                        postal_code=postal_code, 
                        town=town, 
                        city=city, 
                        additional_info=additional_info
                    )
                    print(f"DEBUG - Resultado add_commercial_contact: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_commercial_contact: {str(add_error)}")
                    raise
            elif new_category_index == 2:  # Email no válido
                reason = provided_reason or "Marcado manualmente como inválido"
                try:
                    result = self.db_manager.add_invalid_email(email, name, reason)
                    print(f"DEBUG - Resultado add_invalid_email: {result}")
                except Exception as add_error:
                    print(f"ERROR en add_invalid_email: {str(add_error)}")
                    raise
                
            # Confirmar los cambios
            try:
                self.db_manager.conn.commit()
                print(f"DEBUG - Cambios confirmados en la base de datos")
            except Exception as commit_error:
                print(f"ERROR en commit: {str(commit_error)}")
                raise
                
            return True
            
        except Exception as e:
            print(f"ERROR CRÍTICO al reclasificar el contacto {email}: {str(e)}")
            traceback.print_exc()  # Esto imprimirá el traceback completo
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
                                 ["Nombre", "Email", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"])
        elif current_tab == 1:  # Comerciales
            self._export_category("contactos_comerciales", self.db_manager.get_all_commercial_contacts(),
                                 ["Nombre", "Email", "Empresa", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"])
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
            with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
                # Usar punto y coma como delimitador para mejor compatibilidad con Excel español
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(headers)
                
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_data = []
                for row in data:
                    # Si la fila tiene menos columnas que los encabezados, añadir columnas vacías
                    row_list = list(row)
                    while len(row_list) < len(headers):
                        row_list.append("")
                    # Si la fila tiene más columnas que los encabezados, truncar
                    if len(row_list) > len(headers):
                        row_list = row_list[:len(headers)]
                    formatted_data.append(row_list)
                
                writer.writerows(formatted_data)
            
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
            # Definir headers para cada categoría
            client_headers = ["Nombre", "Email", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"]
            commercial_headers = ["Nombre", "Email", "Empresa", "Fecha Importación", "Código de Cliente", "Dirección", "Código Postal", "Población", "Ciudad", "Información Adicional"]
            invalid_headers = ["Email", "Nombre", "Fecha Importación", "Motivo"]
            
            # Exportar clientes
            clients_path = os.path.join(export_dir, "clientes.csv")
            with open(clients_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(client_headers)
                
                clients = self.db_manager.get_all_clients()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_clients = []
                for row in clients:
                    row_list = list(row)
                    while len(row_list) < len(client_headers):
                        row_list.append("")
                    if len(row_list) > len(client_headers):
                        row_list = row_list[:len(client_headers)]
                    formatted_clients.append(row_list)
                
                writer.writerows(formatted_clients)
            
            # Exportar contactos comerciales
            commercial_path = os.path.join(export_dir, "contactos_comerciales.csv")
            with open(commercial_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(commercial_headers)
                
                commercial_contacts = self.db_manager.get_all_commercial_contacts()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_commercial = []
                for row in commercial_contacts:
                    row_list = list(row)
                    while len(row_list) < len(commercial_headers):
                        row_list.append("")
                    if len(row_list) > len(commercial_headers):
                        row_list = row_list[:len(commercial_headers)]
                    formatted_commercial.append(row_list)
                
                writer.writerows(formatted_commercial)
            
            # Exportar emails no válidos
            invalid_path = os.path.join(export_dir, "emails_no_validos.csv")
            with open(invalid_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                writer.writerow(invalid_headers)
                
                invalid_emails = self.db_manager.get_all_invalid_emails()
                # Asegurar que cada fila tiene el número correcto de columnas
                formatted_invalid = []
                for row in invalid_emails:
                    row_list = list(row)
                    while len(row_list) < len(invalid_headers):
                        row_list.append("")
                    if len(row_list) > len(invalid_headers):
                        row_list = row_list[:len(invalid_headers)]
                    formatted_invalid.append(row_list)
                
                writer.writerows(formatted_invalid)
            
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
