# GigaChat Digest Tool

LangChain-инструмент для генерации корпоративных аналитических **дайджестов**
в формате PowerPoint на основе xlsx-данных и GigaChat.

## Что это

Дайджест — это структурированный обзорный документ в стиле «Голос IT»:

- **Cover-слайд** с заголовком, KPI-карточками и pill-метками источников.
- **Topic-слайды** по продуктам/направлениям — с цитатами сотрудников,
  числом упоминаний, бейджами «new» для новых тем.
- **Градиентный фон** (бирюзовый → персиковый по умолчанию, адаптируется
  под промпт пользователя).

Не презентация со слайдами разных типов, а именно дайджест с фиксированной
структурой: обложка + темы.

## Архитектура

```
xlsx → ExcelReader → DataContext ──┐
                                   ├→ GigaChat → JSON → DigestSpec → DigestBuilder → .pptx
                  style_prompt ────┘                      ↑
                                                          │
                                                 (валидация + санитайзеры)
```

## Установка

```bash
pip install python-pptx pandas openpyxl pydantic lxml \
            langchain-core langchain-gigachat
```

## Быстрый старт

```python
from langchain_gigachat import GigaChat
from pptx_generator import GenerateDigestTool

llm = GigaChat(credentials="...", model="GigaChat-Pro", temperature=0.2)
tool = GenerateDigestTool(llm=llm)

result = tool.invoke({
    "xlsx_path": "/path/to/data.xlsx",
    "style_prompt": "Корпоративный дайджест в пастельной палитре",
    "output_path": "/path/to/digest.pptx",
})
```

## Требования к xlsx

Два листа с точными именами:

- **`Сырье`** — табличные данные. Полезные колонки (необязательны):
  - `Продукт` или `Тема` — для группировки в topic-слайды
  - `Цитата` — для цитат сотрудников в карточках
  - `Источник` — для pill-меток источников
  - `Категория` / `Проблема` — для названий items внутри темы
- **`Суммаризация`** — текст с общим обзором (любой формат).

LLM сам разберётся, как сгруппировать данные в формат дайджеста.

## Структура дайджеста

| Слайд | Содержание |
|---|---|
| **Cover** | Заголовок + подзаголовок + описание + 4 KPI + до 6 pill-меток источников + шапка с датой/периодом |
| **Topic 1..N** | Название темы + 4 KPI по теме + 1-4 проблемы с цитатами/упоминаниями/периодом + бейджи «new» + pill-метки источников темы + футер с метаданными |

## Адаптация стиля

Палитра из 8 цветов подбирается под промпт пользователя:

- `gradient_start`, `gradient_end` — горизонтальный градиент фона
- `card_bg` — фон карточек тем (пастельный)
- `kpi_bg` — фон KPI-карточек
- `text_dark`, `text_muted`, `accent`, `badge` — текст и акценты

Примеры стилевых промптов:
```python
# Дефолтный «Голос IT»
"Корпоративный дайджест в спокойной бирюзово-персиковой палитре"

# Тёмный для совета директоров
"Премиум-дайджест в тёмной палитре глубокого navy с золотыми акцентами"

# Стартап-стиль
"Энергичный дайджест в ярких современных цветах"
```

## Защита от рекомендаций

Дайджест — это **только наблюдения**, не консультация. Защита двухуровневая:

- **В промпте**: запрет на императивный язык («следует», «необходимо»).
- **В Pydantic-валидаторе**: программный санитайзер фильтрует темы и items
  с маркерами рекомендаций.

## Интеграция с твоим LangGraph-агентом

Tool **не перезаписывает** твой системный промпт по умолчанию
(`inject_system_prompt=False`). Если у тебя в LangGraph есть персона
вроде «Привет, ты Агата» — она остаётся работать. Вся инструкция уходит
в одно HumanMessage.

```python
from langgraph.prebuilt import create_react_agent
from pptx_generator import GenerateDigestTool

tool = GenerateDigestTool(llm=llm)
agent = create_react_agent(llm, tools=[*your_tools, tool])
```

## Структура кода

```
src/pptx_generator/
├── schemas.py       — DigestSpec, CoverSlide, TopicSlide + санитайзеры
├── excel_reader.py  — чтение xlsx + профилирование колонок
├── prompts.py       — системный промпт + few-shot пример полного дайджеста
├── builder.py       — рендер .pptx (градиент, pills, KPI, бейджи)
├── tool.py          — LangChain BaseTool
└── __init__.py
```

## Backward compatibility

Старые имена классов сохранены как алиасы:

- `GeneratePresentationTool` = `GenerateDigestTool`
- `GeneratePresentationInput` = `GenerateDigestInput`

Если у тебя в коде ещё старые имена — менять не обязательно.
