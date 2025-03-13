import subprocess
import os
import sys
import shutil

def build_exe():
    print("Comenzando a empaquetar la aplicación...")
    
    # Ruta del script principal
    script_path = "csv_manager_final.py"
    
    # Asegurar que los recursos necesarios existen
    required_files = ["logo.png", "logo.ico"]
    for file in required_files:
        if not os.path.exists(file):
            print(f"Error: El archivo '{file}' no existe en el directorio actual.")
            return
    
    # Limpiar archivos anteriores de compilación si existen
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            print(f"Limpiando directorio {folder}...")
            shutil.rmtree(folder)
    
    # Construir la línea de comando para PyInstaller
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        f"--icon=logo.ico",
        f"--name=CSV_Manager",
        "--clean",
        "--add-data=logo.png{}{}".format(os.pathsep, "."),
        "--add-data=logo.ico{}{}".format(os.pathsep, "."),
    ]
    
    # Agregar el script principal
    cmd.append(script_path)
    
    # Ejecutar PyInstaller
    print(f"Ejecutando comando: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, check=True)
        if result.returncode == 0:
            print("\n¡Compilación exitosa!")
            print("El ejecutable se encuentra en la carpeta 'dist'")
            
            # Copiar el README si existe
            if os.path.exists("readme.md"):
                shutil.copy("readme.md", "dist/")
                print("Se ha copiado el archivo readme.md a la carpeta dist/")
                
            print("\nIMPORTANTE: La base de datos SQLite se creará en el mismo directorio")
            print("donde se ejecute la aplicación, que será dist/CSV_Manager.exe")
        else:
            print(f"\nError al compilar. Código de salida: {result.returncode}")
    except subprocess.CalledProcessError as e:
        print(f"\nError al ejecutar PyInstaller: {e}")
    except Exception as e:
        print(f"\nError inesperado: {e}")

if __name__ == "__main__":
    build_exe() 