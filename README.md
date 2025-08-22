# partners_nat_enric

рограмма для дополнения «Списка для партнёра» колонкой **«новый терминал»** по базе `MID↔TID`.
- Сопоставляет `АЗС` из списка с `MID` из базы.
- Вставляет колонку «новый терминал» и подставляет `TID`.
- Подсвечивает совпавшие ячейки в колонке «АЗС» жёлтым.
- Имя итогового файла включает введённое название партнёра.

## Как запустить на Windows (без установки Python)
Скачайте последний **Windows artifact** из Actions или из Releases, разархивируйте и запустите `PartnerListEnricher.exe`.

## Формат входных файлов
- **База терминалов**: Excel с колонками `MID`, `TID` (регистр не важен, пробелы игнорируются).
- **Список для партнёра**: Excel с колонкой `АЗС` (обязательна). «Терминал»/«Адрес» — необязательны.

## Сборка локально
```powershell
py -m venv .venv
.venv\Scripts\pip install -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller --noconfirm --noconsole --clean ^
  --name PartnerListEnricher ^
  --icon app/icon.ico ^
  --add-data "app/icon.ico;." ^
  app\main.py
