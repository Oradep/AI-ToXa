import subprocess
import json
import requests
import keyboard
import pyperclip
import time
import os
import win32com.client

# === 1. Загружаем system-промпт ===
prompt_path = os.path.join(os.path.dirname(__file__), "prompt.txt")
with open(prompt_path, "r", encoding="utf-8") as f:
    system_prompt = f.read()

# === 2. Базовая конфигурация ===
API_URL = "http://localhost:4891/v1/chat/completions"
HEADERS = {
    "Content-Type": "application/json",
    "Authorization": "Bearer not-needed"
}

print("\nТоха запущен! Пиши свои команды. Для выхода напиши 'выход'.\n")


#обработчики
#НАПИСАНИЕ ТЕКСТА
def type_text_smart(text):
    old_clipboard = pyperclip.paste()
    pyperclip.copy(text)
    time.sleep(0.1)
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.1)
    pyperclip.copy(old_clipboard)


#ЗАПУСК ПРОГРАММ
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







while True:
    user_input = input("Ты: ")
    if user_input.lower() in ["выход", "quit", "exit"]:
        print("Пока! Тоха уходит отдыхать.")
        break

    # === 3. Формируем НОВЫЙ список сообщений каждый раз ===
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_input}
    ]

    # === 4. Отправляем запрос к GPT4All ===
    response = requests.post(
        API_URL,
        headers=HEADERS,
        json={
            "model": "Llama 3 8B Instruct", #Llama 3 8B Instruct; DeepSeek-R1-Distill-Qwen-14B
            "max_tokens": 512,
            "messages": messages
        }
    )

    result = response.json()
    content = result["choices"][0]["message"]["content"]

    print("\n=====\n")
    print(f"Сырые действия от модели:\n{content}")
    print("\n=====\n")

    # === 5. Выполняем все действия по очереди ===
    try:
        actions = json.loads(content)
        if not isinstance(actions, list):
            actions = [actions]

        for action in actions:
            act = action.get("action")
            params = action.get("params", {})

            if act == "say":
                print(f"Тоха говорит: {params.get('response')}")

            elif act == "open_app":
                app = params.get("app")
                print(f"Открываю приложение: {app}")
                open_app(app)

            elif act == "type_text":
                text = params.get("text")
                print(f"Печатаю: {text}")
                time.sleep(1)
                type_text_smart(text)

            elif act == "run_terminal":
                command = params.get("command")
                print(f"Запускаю команду в терминале: {command}")
                subprocess.run(command, shell=True)

            elif act == "none":
                print("Действие: ничего не делать")

            else:
                print(f"Неизвестное действие: {act}")

    except Exception as e:
        print(f"Парсинг/запуск не удался: {e}")

    print("\n=====\n")
