import matplotlib.pyplot as plt
import pandas as pd

# Данные (подогнаны под указанные суммы)
dates = [
    '12.05.2025', '13.05.2025', '14.05.2025', '15.05.2025', '16.05.2025', '17.05.2025', '18.05.2025',
    '19.05.2025', '20.05.2025', '21.05.2025', '22.05.2025', '23.05.2025', '24.05.2025', '25.05.2025'
]
requests = [
    110, 121, 113, 134, 114, 72, 67,  # Умеренный режим (сумма = 731)
    85, 78, 76, 76, 86, 56, 68       # Редкий режим (сумма = 525)
]

# Создаем DataFrame
df = pd.DataFrame({
    'Дата': pd.to_datetime(dates, format='%d.%m.%Y'),
    'Заявки': requests,
    'Режим': ['Умеренный']*7 + ['Редкий']*7
})

# Построение графика
plt.figure(figsize=(14, 7))
bars = plt.bar(df['Дата'], df['Заявки'],
               color=['#4B9CFF' if mode == 'Умеренный' else '#34C759'
                     for mode in df['Режим']],
               edgecolor='black', linewidth=0.5)

# Подписи и оформление
plt.title('Количество заявок на сервис Докма (12.05 – 25.05.2025)', pad=20, fontsize=14, weight='bold')
plt.xlabel('Дата', fontsize=12)
plt.ylabel('Количество заявок', fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Разделительная линия между неделями
plt.axvline(x=pd.to_datetime('18.05.2025', format='%d.%m.%Y') + pd.Timedelta(days=0.5),
            color='gray', linestyle='--', linewidth=1.5, label='Переход режимов')

# Подписи значений над столбцами
for bar in bars:
    height = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2., height + 2,
             f'{int(height)}',
             ha='center', va='bottom', fontsize=10)

# Легенда
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor='#4B9CFF', edgecolor='black', label='Умеренный режим (12.05–18.05)'),
    Patch(facecolor='#34C759', edgecolor='black', label='Редкий режим (19.05–25.05)')
]
plt.legend(handles=legend_elements, loc='upper left', fontsize=10)

plt.tight_layout()
plt.show()