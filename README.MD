# LLM-Parser: Характеристики печатных плат

Проект представляет собой интеллектуальный парсер для извлечения технических характеристик печатных плат из Excel файлов.

Характеристики печатных плат в Excel файлах могут иметь различную структуру, что требует интеллектуального анализа с использованием языковых моделей.

## Project struct
- `main.py` : FastAPI application run file
- **`app`** : Application's modules
    - `utils.py` : Parsing functions
    - `models.py` : Pydantic models for PydanticOutputParser
    - `logger.py` : creating logger
    - `interface.py` : Gradio interface
- **`data`** : examples of PCB characteristics files
- **`media`** : photo and video files for readme
- `requirements.txt` : python requirements

## Run instructions
1. Create venv in project directory `python -m venv venv`
2. Activate venv `venv/Scripts/activate` (windows) or `source venv/bin/activate` (linux)
3. Install requirements.txt `pip install -r requirements.txt`
4. Run application (run in project directory)
    - FastAPI: `fastapi run main.py`
    - Uvicorn: `uvicorn main:app`
5. Open Web-Interface at http://localhost:8000/

## Implemented functionality

### Модуль характеристик печатных плат (Excel)
1. Загрузка Excel файла (.xlsx, .xls) из хранилища
2. Автоматическое извлечение характеристик печатных плат:
   - Название компании
   - Название печатной платы
   - Количество
   - Материал основания
   - Толщина фольги
   - Количество слоев
   - Размер платы
   - Панелизация
3. Предварительный просмотр результатов парсинга
4. Загрузка результатов в файлы .csv, .xlsx, .json

### Интерфейс
- Современный веб-интерфейс для работы с Excel файлами
- Поддержка русского языка
- Интуитивно понятное управление

![](/media/upload.png)
![](/media/res.png)
![](/media/download.png)

## Examples of the parser's work

### Обработка характеристик печатных плат

В проекте представлен пример Excel файла с характеристиками печатных плат для демонстрации функциональности парсера.
