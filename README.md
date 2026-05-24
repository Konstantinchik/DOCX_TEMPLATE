# DOCX_TEMPLATE — InspectionCard Generator

Автоматическая генерация заполненных карт инспекции (DOCX) из данных Excel.

## Быстрый старт

```
1. Положи .docx шаблон в корень этой папки
2. Дважды кликни: vba\LAUNCH_full_pipeline.ps1
3. Открой output\*_Generator.xlsm → Alt+F11 → вставь VBA → Alt+F8 → Run
```

## Структура проекта

```
DOCX_TEMPLATE/
├── README.md                      ← этот файл
├── templates.md                   ← реестр шаблонов + формат map.json
│
├── skills/
│   ├── analyze_docx.py            ← анализ DOCX → history/*_map.json
│   └── generate_xlsm.py           ← генерация xlsm из map.json
│
├── history/
│   └── *_map.json                 ← карты таблиц (авто, не редактировать)
│   └── *_report.md                ← читаемые отчёты об анализе
│
├── tests/
│   ├── test_analyze.py            ← unit-тесты анализатора и генератора
│   └── LAUNCH_tests.ps1           ← запуск тестов
│
├── vba/
│   ├── LAUNCH_analyze.ps1         ← только анализ нового .docx
│   ├── LAUNCH_generate.ps1        ← только генерация xlsm из map.json
│   └── LAUNCH_full_pipeline.ps1   ← analyze + generate за один запуск ★
│
├── output/                        ← сюда кладутся готовые .xlsm и .docx
└── images/                        ← (опц.) Figure1.png, Figure2.png, Figure3.png
```

## Добавление нового шаблона

```powershell
# 1. Положи новый .docx в корень папки
# 2. Запусти:
.\vba\LAUNCH_full_pipeline.ps1
# 3. Откроется output\ с готовым xlsm
```

## Зависимости Python

```powershell
pip install lxml openpyxl
```

## Как работает пайплайн

```
.docx файл
    ↓
analyze_docx.py
    → history/*_map.json   (структура таблиц, строки Mech/Insp, Finding List)
    → history/*_report.md  (читаемый отчёт)
    ↓
generate_xlsm.py
    → output/*_Generator.xlsm
         ├── INSTRUCTIONS  (руководство)
         ├── DATA          (шапка, финальный блок — синие ячейки)
         ├── FINDINGS      (таблица дефектов)
         ├── SIGNERS       (подписи Mech/Insp)
         └── VBA_CODE      (макрос под конкретный .docx)
    ↓
Пользователь:
    Alt+F11 → Module → Paste VBA → Close
    Alt+F8  → FillInspectionCard → Run
    ↓
output\RO_PN_SN.docx  ← заполненная карта
```

## Тесты

```powershell
.\tests\LAUNCH_tests.ps1
# или
python tests\test_analyze.py --docx "myfile.docx" -v
```

## Зарегистрированные шаблоны

| Файл | Статус | Map |
|------|--------|-----|
| JC STR-00324 INSP BALL MAT (D2557216001800) Рев.0.docx | ✅ Готов | history/JC_STR_00324_INSP_BALL_MAT_D2557216001800_Рев0_map.json |
