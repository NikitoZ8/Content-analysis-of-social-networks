from pathlib import Path
import re
import pandas as pd
import matplotlib.pyplot as plt

# =========================
# НАСТРОЙКИ
# =========================
INPUT_FILE = Path("Контент-анализ Мотивация Успех.xlsx")
OUTPUT_DIR = Path("results_content_analysis")
OUTPUT_DIR.mkdir(exist_ok=True)

# Если файл лежит не рядом со скриптом, укажи полный путь, например:
# INPUT_FILE = Path(r"C:\Users\User\Desktop\Контент-анализ Мотивация Успех.xlsx")

MULTI_COLUMNS = [
    "Эмоция",
    "Триггеры",
    "Аргументация",
    "Призыв к действию",
    "Образ успеха",
]
SINGLE_COLUMNS = ["Платформа", "Тип успеха"]

# =========================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================
def clean_value(value):
    """Чистит значение ячейки."""
    if pd.isna(value):
        return "Не указано"
    value = str(value).strip()
    return value if value else "Не указано"


def split_multi(value):
    """Разделяет ячейки, где несколько категорий записаны через запятую."""
    value = clean_value(value)
    parts = [p.strip() for p in value.split(",")]
    return [p for p in parts if p]


def frequency_table(df, column, multi=False):
    """Считает частоту и процент по одному столбцу."""
    if multi:
        values = []
        for item in df[column]:
            values.extend(split_multi(item))
        s = pd.Series(values)
        total = len(df)  # процент от количества постов, а не от количества всех отметок
    else:
        s = df[column].apply(clean_value)
        total = len(s)

    counts = s.value_counts().rename("Количество")
    percent = (counts / total * 100).round(2).rename("Процент")
    return pd.concat([counts, percent], axis=1)


def plot_bar(table, title, filename, ylabel="Количество"):
    """Строит и сохраняет столбчатую диаграмму."""
    plt.figure(figsize=(11, 6))
    table["Количество"].plot(kind="bar")
    plt.title(title)
    plt.ylabel(ylabel)
    plt.xlabel("")
    plt.xticks(rotation=35, ha="right")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=300)
    plt.close()


def plot_pie(table, title, filename):
    """Строит круговую диаграмму."""
    plt.figure(figsize=(8, 8))
    table["Количество"].plot(kind="pie", autopct="%1.1f%%")
    plt.title(title)
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=300)
    plt.close()


def explode_column(df, column):
    """Разворачивает мультикатегориальный столбец в несколько строк."""
    temp = df.copy()
    temp[column] = temp[column].apply(split_multi)
    return temp.explode(column)


def plot_grouped_percent(crosstab, title, filename):
    """Строит grouped bar chart по процентам."""
    ax = crosstab.plot(kind="bar", figsize=(12, 6))
    plt.title(title)
    plt.ylabel("% внутри платформы")
    plt.xlabel("")
    plt.xticks(rotation=0)
    plt.legend(title="Категория", bbox_to_anchor=(1.02, 1), loc="upper left")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=300)
    plt.close()


def plot_heatmap(matrix, title, filename):
    """Простая тепловая карта без seaborn."""
    plt.figure(figsize=(10, 6))
    plt.imshow(matrix.values, aspect="auto")
    plt.colorbar(label="%")
    plt.title(title)
    plt.xticks(range(len(matrix.columns)), matrix.columns, rotation=35, ha="right")
    plt.yticks(range(len(matrix.index)), matrix.index)

    # подписи в ячейках
    for i in range(len(matrix.index)):
        for j in range(len(matrix.columns)):
            plt.text(j, i, f"{matrix.iloc[i, j]:.1f}", ha="center", va="center")

    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / filename, dpi=300)
    plt.close()

# =========================
# ЗАГРУЗКА ДАННЫХ
# =========================
df = pd.read_excel(INPUT_FILE)
df.columns = [str(c).strip() for c in df.columns]

# Приводим нужные столбцы к строковому виду
for col in SINGLE_COLUMNS + MULTI_COLUMNS:
    if col in df.columns:
        df[col] = df[col].apply(clean_value)

# Технические признаки для дополнительного анализа
if "Текст" in df.columns:
    df["Длина текста"] = df["Текст"].fillna("").astype(str).apply(len)
else:
    df["Длина текста"] = 0

# =========================
# 1. ОБЩИЕ ЧАСТОТЫ
# =========================
all_stats = {}
for col in SINGLE_COLUMNS:
    if col in df.columns:
        all_stats[col] = frequency_table(df, col, multi=False)
        all_stats[col].to_csv(OUTPUT_DIR / f"frequency_{col}.csv", encoding="utf-8-sig")
        plot_bar(all_stats[col], f"Распределение: {col}", f"bar_{col}.png")
        if col == "Тип успеха":
            plot_pie(all_stats[col], "Тип успеха, %", "pie_Тип_успеха.png")

for col in MULTI_COLUMNS:
    if col in df.columns:
        all_stats[col] = frequency_table(df, col, multi=True)
        all_stats[col].to_csv(OUTPUT_DIR / f"frequency_{col}.csv", encoding="utf-8-sig")
        plot_bar(all_stats[col], f"Частота категорий: {col}", f"bar_{col}.png")

# =========================
# 2. СРАВНЕНИЕ TG И TIKTOK
# =========================
for col in ["Тип успеха"]:
    if "Платформа" in df.columns and col in df.columns:
        platform_ct = pd.crosstab(df["Платформа"], df[col], normalize="index") * 100
        platform_ct = platform_ct.round(2)
        platform_ct.to_csv(OUTPUT_DIR / f"platform_compare_{col}.csv", encoding="utf-8-sig")
        plot_grouped_percent(platform_ct, f"Сравнение платформ по параметру: {col}", f"platform_compare_{col}.png")

for col in MULTI_COLUMNS:
    if "Платформа" in df.columns and col in df.columns:
        ex = explode_column(df, col)
        platform_ct = pd.crosstab(ex["Платформа"], ex[col], normalize="index") * 100
        platform_ct = platform_ct.round(2)
        platform_ct.to_csv(OUTPUT_DIR / f"platform_compare_{col}.csv", encoding="utf-8-sig")
        plot_grouped_percent(platform_ct, f"Сравнение платформ по параметру: {col}", f"platform_compare_{col}.png")

# =========================
# 3. СВЯЗИ МЕЖДУ ПРИЗНАКАМИ
# =========================
# Связь: тип успеха × триггеры
if "Тип успеха" in df.columns and "Триггеры" in df.columns:
    trig = explode_column(df, "Триггеры")
    matrix = pd.crosstab(trig["Тип успеха"], trig["Триггеры"], normalize="index") * 100
    matrix = matrix.round(2)
    matrix.to_csv(OUTPUT_DIR / "relation_Тип_успеха_Триггеры.csv", encoding="utf-8-sig")
    plot_heatmap(matrix, "Связь: тип успеха × триггеры", "heatmap_Тип_успеха_Триггеры.png")

# Связь: платформа × образ успеха
if "Платформа" in df.columns and "Образ успеха" in df.columns:
    success = explode_column(df, "Образ успеха")
    matrix = pd.crosstab(success["Платформа"], success["Образ успеха"], normalize="index") * 100
    matrix = matrix.round(2)
    matrix.to_csv(OUTPUT_DIR / "relation_Платформа_Образ_успеха.csv", encoding="utf-8-sig")
    plot_heatmap(matrix, "Связь: платформа × образ успеха", "heatmap_Платформа_Образ_успеха.png")

# Связь: платформа × эмоция
if "Платформа" in df.columns and "Эмоция" in df.columns:
    emo = explode_column(df, "Эмоция")
    matrix = pd.crosstab(emo["Платформа"], emo["Эмоция"], normalize="index") * 100
    matrix = matrix.round(2)
    matrix.to_csv(OUTPUT_DIR / "relation_Платформа_Эмоция.csv", encoding="utf-8-sig")
    plot_heatmap(matrix, "Связь: платформа × эмоция", "heatmap_Платформа_Эмоция.png")

# =========================
# 4. АНАЛИЗ ПО КАНАЛАМ
# =========================
if "Название канала" in df.columns:
    # средняя длина текста по каналам
    text_len = df.groupby(["Платформа", "Название канала"], as_index=False)["Длина текста"].mean()
    text_len["Длина текста"] = text_len["Длина текста"].round(1)
    text_len.to_csv(OUTPUT_DIR / "avg_text_length_by_channel.csv", index=False, encoding="utf-8-sig")

    plt.figure(figsize=(12, 7))
    text_len.sort_values("Длина текста").set_index("Название канала")["Длина текста"].plot(kind="barh")
    plt.title("Средняя длина текста по каналам")
    plt.xlabel("Среднее количество символов")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig(OUTPUT_DIR / "avg_text_length_by_channel.png", dpi=300)
    plt.close()

    # доля долгосрочного успеха по каналам
    if "Тип успеха" in df.columns:
        long_success = df.assign(is_long=df["Тип успеха"].eq("Долгосрочный"))
        long_rate = long_success.groupby(["Платформа", "Название канала"], as_index=False)["is_long"].mean()
        long_rate["Доля долгосрочного успеха, %"] = (long_rate["is_long"] * 100).round(2)
        long_rate = long_rate.drop(columns=["is_long"])
        long_rate.to_csv(OUTPUT_DIR / "long_success_rate_by_channel.csv", index=False, encoding="utf-8-sig")

        plt.figure(figsize=(12, 7))
        long_rate.sort_values("Доля долгосрочного успеха, %").set_index("Название канала")["Доля долгосрочного успеха, %"].plot(kind="barh")
        plt.title("Доля долгосрочного успеха по каналам")
        plt.xlabel("% постов")
        plt.ylabel("")
        plt.tight_layout()
        plt.savefig(OUTPUT_DIR / "long_success_rate_by_channel.png", dpi=300)
        plt.close()

# =========================
# 5. ИТОГОВЫЙ ОТЧЕТ В EXCEL
# =========================
report_path = OUTPUT_DIR / "summary_tables.xlsx"
with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
    for name, table in all_stats.items():
        safe_name = re.sub(r"[\\/*?:\[\]]", "_", name)[:31]
        table.to_excel(writer, sheet_name=safe_name)

    if "Платформа" in df.columns and "Тип успеха" in df.columns:
        pd.crosstab(df["Платформа"], df["Тип успеха"], normalize="index").mul(100).round(2).to_excel(
            writer, sheet_name="Платформа_Тип успеха"
        )

print("Анализ завершён.")
print(f"Папка с результатами: {OUTPUT_DIR.resolve()}")
print("Созданы PNG-графики, CSV-таблицы и Excel-файл summary_tables.xlsx")
