### Basement

Небольшое настольное приложение на PyQt5 для конвертации табличного отчёта (.txt) в оформленные Excel‑файлы (.xlsx), а также для экспресс‑ и полосового анализа данных. Есть две вкладки: «Тензометрия» и «Полосовой анализ». Отдельный скрипт `addition.py` строит гистограммы по готовому `.xlsx`.

### Структура проекта
- `main.py` — GUI и основная логика (PyQt5, openpyxl, chardet). Загружает ресурсы (`myicon.ico`, `header.png`, `dinosaur.gif`) через функцию `getIconPath()`:
  - при обычном запуске — из текущей папки (`.`)
  - в сборке PyInstaller — из каталога `sys._MEIPASS`
- `addition.py` — построение гистограмм (openpyxl.chart) по данным из Excel.
- `header.png`, `dinosaur.gif`, `myicon.ico` — ресурсы интерфейса.

### Установка
```powershell
cd E:\Max\Basement
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
```

### Запуск из исходников
```powershell
py main.py
```

### Сборка в .exe (Windows, PyInstaller)
- Сборка в один файл (.exe):
```powershell
py -m PyInstaller --noconsole --onefile --name Basement --icon myicon.ico \
  --add-data "myicon.ico;." \
  --add-data "header.png;." \
  --add-data "dinosaur.gif;." \
  main.py
```

- Альтернатива (папка приложения):
```powershell
py -m PyInstaller --noconsole --name Basement --icon myicon.ico \
  --add-data "myicon.ico;." \
  --add-data "header.png;." \
  --add-data "dinosaur.gif;." \
  main.py
```

- Где искать результат:
  - onefile: `dist/Basement.exe`
  - папка: `dist/Basement/Basement.exe`

### Примечания
- Для Windows в `--add-data` используется разделитель `;` (как в примерах). На macOS/Linux — `:`.
- При сборке в один файл первый запуск может быть дольше из‑за распаковки во временную папку.
- Ресурсы будут найдены благодаря `getIconPath()` в `main.py`.