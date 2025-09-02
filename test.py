import os
import subprocess
import win32com.client

def find_shortcut_path(app_name):
    # Пути к меню Пуск (общие и пользователя)
    start_menu_paths = [
        os.environ["PROGRAMDATA"] + r"\Microsoft\Windows\Start Menu\Programs",
        os.environ["APPDATA"] + r"\Microsoft\Windows\Start Menu\Programs"
    ]
    
    for base_path in start_menu_paths:
        for root, dirs, files in os.walk(base_path):
            for file in files:
                if file.lower().endswith(".lnk") and app_name.lower() in file.lower():
                    return os.path.join(root, file)
    return None

def open_app(app_name):
    # Если это Steam-ссылка
    if app_name.startswith("steam://"):
        try:
            print(f"Запускаю Steam-игру: {app_name}")
            subprocess.Popen(["cmd", "/c", "start", "", app_name], shell=True)
        except Exception as e:
            print(f"Ошибка запуска Steam-ссылки: {e}")
        return

    # Поиск по ярлыкам
    shortcut_path = find_shortcut_path(app_name)
    if shortcut_path:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            target = shortcut.TargetPath
            print(f"Открываю: {target}")
            subprocess.Popen(target)
        except Exception as e:
            print(f"Ошибка запуска приложения: {e}")
    else:
        print(f"Приложение '{app_name}' не найдено.")

        
# Пример использования
open_app("Блокнот")  # или "Notepad", "Telegram", "Discord" и т.п.
