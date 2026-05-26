"""
Полный пример использования: GigaChat + tool + xlsx → pptx.

Запуск:
    pip install python-pptx pandas openpyxl pydantic langchain-core langchain-gigachat
    export GIGACHAT_CREDENTIALS="ваш_ключ"
    python example_usage.py
"""
import os
from pathlib import Path

import pandas as pd

from pptx_generator import GeneratePresentationTool


# --------------------------------------------------------------------------- #
# 1. Готовим демонстрационный xlsx
# --------------------------------------------------------------------------- #

def create_demo_xlsx(path: Path) -> Path:
    """Создаёт показательный xlsx с двумя нужными листами."""
    raw_df = pd.DataFrame({
        "Регион": ["Москва", "СПб", "Казань", "Москва", "СПб", "Казань",
                   "Москва", "СПб", "Казань", "Новосибирск"],
        "Месяц": ["Янв", "Янв", "Янв", "Фев", "Фев", "Фев",
                  "Мар", "Мар", "Мар", "Мар"],
        "Выручка_млн": [120, 85, 43, 135, 92, 48, 142, 88, 52, 38],
        "Жалобы": [12, 8, 4, 14, 9, 5, 17, 10, 6, 7],
        "NPS": [67, 71, 75, 62, 68, 73, 58, 65, 72, 70],
        "Категория_проблем": [
            "Скорость", "Качество", "Персонал",
            "Скорость", "Скорость", "Качество",
            "Скорость", "Системы", "Качество", "Скорость",
        ],
    })

    summary_df = pd.DataFrame([
        ["Основные проблемы квартала:"],
        [""],
        ["1. Падение NPS на 9 пунктов за квартал — с 67 до 58 в Москве."],
        ["2. Рост жалоб на скорость обслуживания: 42% всех обращений."],
        ["3. Региональная сеть растёт по выручке, но проседает по качеству."],
        ["4. Корреляция: чем выше время ответа, тем ниже оценка клиента."],
        ["5. Декабрь-март — критический период для всех точек."],
        [""],
        ["Рекомендации:"],
        ["- Внедрить SLA на скорость реакции (целевое значение: 4 часа)."],
        ["- Запустить программу обучения для региональных команд."],
        ["- Создать дашборд жалоб с автоматическими алертами."],
    ])

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Сырье", index=False)
        summary_df.to_excel(writer, sheet_name="Суммаризация", index=False, header=False)

    return path


# --------------------------------------------------------------------------- #
# 2. Запускаем tool с GigaChat
# --------------------------------------------------------------------------- #

def main():
    # Подготовка данных
    xlsx_path = Path("/tmp/demo_data.xlsx")
    create_demo_xlsx(xlsx_path)
    print(f"Создан тестовый xlsx: {xlsx_path}")

    # GigaChat через LangChain
    # pip install langchain-gigachat
    from langchain_gigachat import GigaChat

    llm = GigaChat(
        credentials=os.environ["GIGACHAT_CREDENTIALS"],
        model="GigaChat-Pro",      # или GigaChat-Max для лучших результатов
        verify_ssl_certs=False,    # для корпоративной сети может потребоваться
        temperature=0.3,           # пониже — для стабильного JSON
        scope="GIGACHAT_API_PERS",
    )

    # Инструмент
    tool = GeneratePresentationTool(llm=llm, max_retries=3)

    # Вариант 1: прямой вызов
    result_path = tool.invoke({
        "xlsx_path": str(xlsx_path),
        "style_prompt": (
            "Презентация для совета директоров банка. Премиальный корпоративный "
            "стиль, тёмная палитра с акцентом на тревожности проблем. "
            "Шрифт — Georgia для заголовков. Не больше 12 слайдов. Сфокусироваться "
            "на бизнес-импакте, добавить graphics по динамике жалоб и структуре проблем."
        ),
        "output_path": "/tmp/quarterly_report.pptx",
    })
    print(f"Готово: {result_path}")


# --------------------------------------------------------------------------- #
# 3. Использование в составе агента
# --------------------------------------------------------------------------- #

def example_with_agent():
    """Пример: tool как один из доступных агенту инструментов."""
    from langchain_gigachat import GigaChat
    from langchain.agents import AgentExecutor, create_tool_calling_agent
    from langchain_core.prompts import ChatPromptTemplate

    llm = GigaChat(credentials=os.environ["GIGACHAT_CREDENTIALS"], model="GigaChat-Pro")

    tool = GeneratePresentationTool(llm=llm)

    prompt = ChatPromptTemplate.from_messages([
        ("system", "Ты — помощник аналитика. Используй доступные инструменты."),
        ("human", "{input}"),
        ("placeholder", "{agent_scratchpad}"),
    ])

    agent = create_tool_calling_agent(llm, [tool], prompt)
    executor = AgentExecutor(agent=agent, tools=[tool], verbose=True)

    executor.invoke({
        "input": (
            "Сгенерируй презентацию по файлу /tmp/demo_data.xlsx. "
            "Стиль — современный технологический, для презентации продукта."
        )
    })


if __name__ == "__main__":
    main()
