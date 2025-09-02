import subprocess
import json
import requests
import keyboard
import pyperclip
import time
import os
import win32com.client
import speech_recognition as sr
import pyttsx3
import webbrowser
import ctypes



# === 1. Загрузка system-промпта ===
prompt_path = os.path.join(os.path.dirname(__file__), "prompt.txt")
with open(prompt_path, "r", encoding="utf-8") as f:
    system_prompt = f.read()

# === 2. Конфигурация API ===
API_URL = "http://localhost:4891/v1/chat/completions"
HEADERS = {
    "Content-Type": "application/json",
    "Authorization": "Bearer not-needed"
}

# === 3. Голосовой движок и распознавание ===
r = sr.Recognizer()
engine = pyttsx3.init(driverName='sapi5')  # Для Windows
voices = engine.getProperty('voices')

for voice in voices:
    if "Irina" in voice.name or "Russian" in voice.name:
        engine.setProperty('voice', voice.id)
        break

engine.setProperty('volume', 1.0)
engine.setProperty('rate', 170)

def say(text):
    print(f"Тоха говорит: {text}")
    engine.say(text)
    engine.runAndWait()

def listen_for_command_with_wake_word(wake_word="тоха"):
    with sr.Microphone() as source:
        print("Слушаю (ключевое слово + команда)...")

        # Настройка распознавания под микрофон
        r.adjust_for_ambient_noise(source, duration=1)  # автопорог
        r.pause_threshold = 1.5  # тишина, после которой завершается ввод
        #r.energy_threshold = 300  # можно оставить авто, но фикс иногда лучше

        try:
            audio = r.listen(source, timeout=None)  # слушаем до тишины
            phrase = r.recognize_google(audio, language="ru-RU").lower()
            print(f"Распознано: {phrase}")

            if wake_word in phrase:
                idx = phrase.find(wake_word)
                command = phrase[idx + len(wake_word):].strip()
                return command if command else None
        except sr.UnknownValueError:
            print("Не удалось распознать речь")
        except sr.RequestError as e:
            print(f"Ошибка сервиса распознавания: {e}")
    return None



# === 4. Действия ассистента ===
def type_text_smart(text):
    old_clipboard = pyperclip.paste()
    pyperclip.copy(text)
    time.sleep(0.1)
    keyboard.press_and_release('ctrl+v')
    time.sleep(0.1)
    pyperclip.copy(old_clipboard)

def find_shortcut_path(app_name):
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
    if app_name.startswith("steam://"):
        try:
            subprocess.Popen(["cmd", "/c", "start", "", app_name], shell=True)
        except Exception as e:
            say(f"Ошибка запуска Steam: {e}")
        return
    shortcut_path = find_shortcut_path(app_name)
    if shortcut_path:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(shortcut_path)
            target = shortcut.TargetPath
            subprocess.Popen(target)
        except Exception as e:
            say(f"Ошибка запуска: {e}")
    else:
        say(f"Приложение '{app_name}' не найдено")

# === 5. Основной цикл ===
print("\nТоха запущен! Говори 'Тоха' для активации.\n")

while True:
    user_input = listen_for_command_with_wake_word("тоха")
    if not user_input:
        continue

    print(f"Ты (голосом): {user_input}")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_input}
    ]

    try:
        response = requests.post(
            API_URL,
            headers=HEADERS,
            json={
                "model": "Llama 3 8B Instruct",
                "max_tokens": 512,
                "messages": messages
            }
        )

        result = response.json()
        content = result["choices"][0]["message"]["content"]
        print(f"\nОтвет модели:\n{content}\n")

        actions = json.loads(content)
        if not isinstance(actions, list):
            actions = [actions]

        for action in actions:
            act = action.get("action")
            params = action.get("params", {})

            if act == "say":
                say(params.get("response", ""))

            elif act == "open_app":
                app = params.get("app")
                say(f"Открываю {app}")
                open_app(app)

            elif act == "type_text":
                text = params.get("text")
                say("Печатаю текст")
                time.sleep(1)
                type_text_smart(text)

            elif act == "run_terminal":
                command = params.get("command")
                say(f"Выполняю команду")
                subprocess.run(command, shell=True)

            elif act == "open_website":
                url = params.get("url")
                if url:
                    say(f"Открываю сайт")
                    if not url.startswith("http"):
                        url = "https://" + url
                    webbrowser.open(url)
                else:
                    say("Сайт не указан")

            elif act == "system_action":
                action_type = params.get("type")
                if action_type == "shutdown":
                    say("Выключаю компьютер")
                    subprocess.run("shutdown /s /t 1", shell=True)
                elif action_type == "restart":
                    say("Перезагружаю компьютер")
                    subprocess.run("shutdown /r /t 1", shell=True)
                elif action_type == "lock":
                    say("Блокирую компьютер")
                    ctypes.windll.user32.LockWorkStation()
                else:
                    say("Неизвестное системное действие")


            elif act == "none":
                print("Ничего не делать")

            else:
                print(f"Неизвестное действие: {act}")

    except Exception as e:
        print(f"Ошибка при обработке: {e}")
        say("Произошла ошибка")
