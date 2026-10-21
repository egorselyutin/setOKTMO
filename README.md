# Простановка кодов ОКТМО (Streamlit)

Запуск:
    pip install -r requirements.txt
    streamlit run app.py

Рядом с app.py должны лежать: НП_ОКТМО_01012025.xlsx (справочник), win-xp-sound.wav (необязательно),
settings.py, string_lib.py, remover.py, stringsCleaner.py, pipeline.py.
Число процессов: переменная окружения OKTMO_WORKERS (по умолчанию min(CPU, 8)).
Путь к справочнику: переменная окружения OKTMO_FILE.
