import pandas as pd
import re
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO
import xlsxwriter

# Загрузка данных
input_file = "chats---2025-04-01---2025-04-30.xlsx"
try:
    df = pd.read_excel(input_file, header=None)
except FileNotFoundError:
    print(f"Ошибка: Файл {input_file} не найден")
    exit(1)

# Переименование столбцов
df.columns = ["raw_data"]

# Парсинг строк
def parse_row(row):
    try:
        if isinstance(row, str):
            # Улучшенное регулярное выражение для разделения
            match = re.match(r"\[(.*?)\](.*)", row)
            if match:
                timestamp = match.group(1).strip()
                content = match.group(2).strip()
                return {"timestamp": timestamp, "content": content}
    except Exception as e:
        print(f"Ошибка парсинга строки: {row}, ошибка: {str(e)}")
    return {"timestamp": None, "content": str(row)}

# Применение парсинга
parsed = df["raw_data"].apply(parse_row)
chat_df = pd.DataFrame(parsed.tolist())

# Удаление пустых строк
chat_df = chat_df.dropna(subset=["content"])
chat_df = chat_df[chat_df["content"].str.strip() != ""]

# Определение типа запроса
def classify_query(text):
    if not isinstance(text, str):
        return "Прочее"
    text = text.lower()
    if any(word in text for word in ["врач", "специалист"]):
        return "Помощь с выбором врача"
    elif any(word in text for word in ["оплата", "цена", "стоимость"]):
        return "Оплата / цена"
    elif any(word in text for word in ["письмо", "почта", "логин", "пароль"]):
        return "Технические вопросы"
    elif any(word in text for word in ["анализ", "результат"]):
        return "Расшифровка анализов"
    elif any(word in text for word in ["спасибо", "хорошо"]):
        return "Обратная связь"
    return "Прочее"

chat_df["query_type"] = chat_df["content"].apply(classify_query)

# Подсчёт времени ответа
chat_df["datetime"] = pd.to_datetime(chat_df["timestamp"], errors='coerce')
chat_df["is_operator"] = chat_df["content"].str.contains("Оператор|администратор", case=False, na=False)
chat_df["is_user"] = ~chat_df["is_operator"]
chat_df["is_closed"] = chat_df["content"].str.contains("закрыл окно чата", case=False, na=False)

# Сохранение в Excel
output_file = "docma_chats_analysis_2025-04.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Лист 1: Чаты
    chat_df.to_excel(writer, sheet_name="Чаты", index=False)

    # Лист 2: Аналитика
    # Расчет среднего времени ответа только для последовательных сообщений
    time_diffs = chat_df[chat_df["is_operator"]]["datetime"].diff()
    mean_response_time = time_diffs[time_diffs.notna()].dt.total_seconds().mean() / 60 if time_diffs.notna().any() else 0

    summary = {
        "Метрика": [
            "Общее количество чатов",
            "Успешные чаты (ответ дан)",
            "Среднее время ответа (мин)",
            "Самый частый тип запроса",
            "Процент «холодных» выходов"
        ],
        "Значение": [
            len(chat_df),
            chat_df["is_operator"].sum(),
            round(mean_response_time, 2) if mean_response_time else "N/A",
            chat_df["query_type"].mode()[0] if not chat_df["query_type"].empty else "N/A",
            round((chat_df["is_closed"].sum() / len(chat_df)) * 100, 2) if len(chat_df) > 0 else 0
        ]
    }
    summary_df = pd.DataFrame(summary)
    summary_df.to_excel(writer, sheet_name="Аналитика", index=False)

    # Лист 3: Диаграммы
    workbook = writer.book
    worksheet = workbook.add_worksheet("Диаграммы")

    # Распределение типов запросов
    query_counts = chat_df["query_type"].value_counts()
    if not query_counts.empty:
        fig, ax = plt.subplots(figsize=(8, 5))
        query_counts.plot(kind="pie", autopct='%1.1f%%', ax=ax, title="Распределение типов запросов")
        plt.tight_layout()

        chart_data = BytesIO()
        plt.savefig(chart_data, format='png')
        plt.close()
        worksheet.insert_image('A1', 'query_distribution.png', {'image_data': chart_data})

    # График активности по времени
    chat_df["hour"] = chat_df["datetime"].dt.hour
    hourly_activity = chat_df.groupby("hour").size()
    if not hourly_activity.empty:
        fig, ax = plt.subplots(figsize=(10, 5))
        hourly_activity.plot(kind="line", marker="o", ax=ax, title="Активность пользователей по часам")
        plt.xlabel("Час")
        plt.ylabel("Количество сообщений")
        plt.grid(True)

        chart_data_hourly = BytesIO()
        plt.savefig(chart_data_hourly, format='png')
        plt.close()
        worksheet.insert_image('A20', 'hourly_activity.png', {'image_data': chart_data_hourly})

print(f"Файл сохранён как {output_file}")