import pandas as pd

# Настройки отображения DataFrame
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# Загрузка данных из Excel
df = pd.read_excel(r'C:\Users\user\Downloads\ExpotCatalog (5).xls')

# Отсортировать DataFrame по столбцу 'Количество просмотров' в порядке убывания
df = df.sort_values(by='Количество просмотров ( не трогать )', ascending=False)

# Удаление строк, где "Количество покупок (не трогать)" меньше 1
df = df[df['Количество покупок ( не трогать )'] >= 1]

# Удаление строк в столбце "Кол-во в МСК" меньше 2
df = df[df['Кол-во в МСК'] >= 2]


# Функция для расчета коэффициента покупок
def get_purchase_coeff(purchases):
    return 0.03 * purchases


# Функция для расчета коэффициента просмотров
def get_view_coeff(views):
    return 0.0000125 * views


# Функция для расчета коэффициента запаса МСК
def get_stock_coeff(purchases):
    return 0.0015 * purchases


# Добавление столбца "Популярность"
df['Популярность'] = df['Количество покупок ( не трогать )'].apply(get_purchase_coeff) + df[
    'Количество просмотров ( не трогать )'].apply(get_view_coeff)

# Добавление столбца "Запас МСК"
df['Запас МСК'] = df['Количество покупок ( не трогать )'].apply(get_stock_coeff)

# Добавление столбца "ИТОГ"
df['ИТОГ'] = df['Популярность'] + df['Запас МСК']

# Сортировка DataFrame по столбцу "ИТОГ" в порядке убывания
df = df.sort_values(by='ИТОГ', ascending=False)

# Добавление столбца "Порядковый номер" в порядке убывания
df['Порядковый номер'] = range(len(df), 0, -1)

# Сохранение отсортированных данных в Excel-файл
df.to_excel('ExpotCatalog_sorted.xlsx', index=False)

# Вывод отсортированного DataFrame
print(df)
