# -*- coding: utf-8 -*-
"""
Контент-анализ мотивационных каналов TG и TikTok.
Скрипт читает Excel-файл, строит графики и сохраняет таблицы результатов.

Как использовать:
1. Положите файл Excel рядом со скриптом или укажите путь в FILE_PATH.
2. Установите библиотеки:
   pip install pandas matplotlib seaborn openpyxl numpy
3. Запустите:
   python analyze_content_final.py

Результаты сохраняются в папку content_analysis_output.
"""

import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns


# =========================
# 0. НАСТРОЙКИ
# =========================
FILE_PATH = "Контент-анализ Мотивация Успех.xlsx"
OUTPUT_DIR = Path("content_analysis_output")
GRAPHS_DIR = OUTPUT_DIR / "graphs"
TABLES_DIR = OUTPUT_DIR / "tables"

for folder in [OUTPUT_DIR, GRAPHS_DIR, TABLES_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.titlesize"] = 14
plt.rcParams["axes.labelsize"] = 11
plt.rcParams["xtick.labelsize"] = 10
plt.rcParams["ytick.labelsize"] = 10


# =========================
# 1. ЗАГРУЗКА И ПОДГОТОВКА ДАННЫХ
# =========================
def normalize_platform(value: object) -> str:
    """Приводит разные варианты названия платформы к TG / TikTok."""
    x = str(value).strip().lower()
    x = x.replace(" ", "")

    tg_values = {"tg", "telegram", "телеграм", "телеграмм", "тг"}
    tiktok_values = {"tiktok", "tik-tok", "tik_tok", "тик-ток", "тикток", "тик_ток", "tt", "тт"}

    if x in tg_values:
        return "TG"
    if x in tiktok_values:
        return "TikTok"
    return str(value).strip()


def clean_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns={
        "Количество подписчиков": "Подписчики",
    })
    return df


def to_number(value: object) -> float:
    """Переводит подписчиков в число: '31 311', '31,311', '31311' -> 31311."""
    if pd.isna(value):
        return np.nan
    x = str(value)
    x = x.replace(" ", "").replace("\xa0", "")
    x = re.sub(r"[^0-9,\.]", "", x)
    if x.count(",") == 1 and x.count(".") == 0:
        x = x.replace(",", ".")
    elif x.count(",") > 0 and x.count(".") > 0:
        x = x.replace(",", "")
    try:
        return float(x)
    except ValueError:
        return np.nan


def split_multi_values(value: object) -> list[str]:
    """
    Разбивает ячейку с несколькими вариантами.
    Например: 'Негатив, Позитив' -> ['Негатив', 'Позитив'].
    """
    if pd.isna(value):
        return []

    text = str(value).strip()
    if not text:
        return []

    text = text.replace(";", ",")
    text = text.replace("\n", ",")
    parts = [p.strip() for p in text.split(",")]
    parts = [p for p in parts if p]

    # небольшая нормализация частых вариантов
    replacements = {
        "Отсутствует": "Нет",
        "Отсутствуют": "Нет",
        "Не определён": "Не определен",
        "Неопределен": "Не определен",
        "Неопределён": "Не определен",
        "Не указан": "Не указан",
    }
    parts = [replacements.get(p, p) for p in parts]
    return parts


def explode_category(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """Создает датафрейм: одна строка = одно значение смысловой единицы."""
    rows = []
    for idx, row in df.iterrows():
        values = split_multi_values(row.get(column))
        for value in values:
            rows.append({
                "Платформа": row.get("Платформа"),
                "Название канала": row.get("Название канала"),
                "Подписчики": row.get("Подписчики"),
                "Значение": value,
            })
    return pd.DataFrame(rows)


def require_columns(df: pd.DataFrame, columns: list[str]) -> None:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(
            "В Excel не найдены нужные столбцы: " + ", ".join(missing) +
            "\nПроверьте названия столбцов в файле."
        )


# загрузка
file_path = Path(FILE_PATH)
if not file_path.exists():
    # запасной вариант: если скрипт запущен из другой папки
    alt = Path.cwd() / FILE_PATH
    if alt.exists():
        file_path = alt

if not file_path.exists():
    raise FileNotFoundError(f"Файл не найден: {FILE_PATH}")

df = pd.read_excel(file_path)
df = clean_column_names(df)

semantic_columns = [
    "Тип успеха",
    "Эмоция",
    "Триггеры",
    "Аргументация",
    "Призыв к действию",
    "Образ успеха",
]

base_columns = ["Платформа", "Название канала", "Подписчики", "Текст"]
require_columns(df, base_columns + semantic_columns)

df["Платформа"] = df["Платформа"].apply(normalize_platform)
df["Подписчики"] = df["Подписчики"].apply(to_number)
df["Длина текста"] = df["Текст"].fillna("").astype(str).str.len()

# Для манипулятивных триггеров условно считаем: Срочность, Страх упустить, Кейс, Убеждение.
# Если в ячейке только 'Нет', значит триггеров нет.
def has_trigger(value: object) -> int:
    values = split_multi_values(value)
    if not values or values == ["Нет"]:
        return 0
    return int(any(v != "Нет" for v in values))


def has_direct_cta(value: object) -> int:
    values = split_multi_values(value)
    direct = {"Купить", "Подписаться", "Перейти"}
    return int(any(v in direct for v in values))


df["Есть триггер"] = df["Триггеры"].apply(has_trigger)
df["Есть CTA"] = df["Призыв к действию"].apply(has_direct_cta)


# =========================
# 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ГРАФИКОВ
# =========================
def save_plot(filename: str) -> None:
    plt.tight_layout()
    plt.savefig(GRAPHS_DIR / filename, dpi=300, bbox_inches="tight")
    plt.close()


def add_bar_labels(ax, fmt="{:.0f}"):
    """Подписывает значения над столбиками."""
    for container in ax.containers:
        labels = []
        for bar in container:
            height = bar.get_height()
            if pd.isna(height) or height == 0:
                labels.append("")
            else:
                labels.append(fmt.format(height))
        ax.bar_label(container, labels=labels, fontsize=8, padding=2)


def plot_platform_distribution():
    counts = df["Платформа"].value_counts().reindex(["TG", "TikTok"]).dropna()

    fig, ax = plt.subplots(figsize=(8, 6))
    counts.plot(kind="bar", ax=ax)
    ax.set_title("01 Распределение единиц анализа по платформам")
    ax.set_xlabel("")
    ax.set_ylabel("Количество единиц анализа")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
    add_bar_labels(ax)
    save_plot("01_распределение_платформ.png")

    counts.to_csv(TABLES_DIR / "01_распределение_платформ.csv", encoding="utf-8-sig")


def plot_frequency_by_platform(column: str, number: str):
    """
    Частотный график для смысловой единицы:
    общее + TG + TikTok.
    """
    exploded_all = explode_category(df, column)
    exploded_tg = explode_category(df[df["Платформа"] == "TG"], column)
    exploded_tiktok = explode_category(df[df["Платформа"] == "TikTok"], column)

    all_counts = exploded_all["Значение"].value_counts()
    tg_counts = exploded_tg["Значение"].value_counts()
    tiktok_counts = exploded_tiktok["Значение"].value_counts()

    combined = pd.DataFrame({
        "Общее": all_counts,
        "TG": tg_counts,
        "TikTok": tiktok_counts,
    }).fillna(0)

    combined = combined.sort_values("Общее", ascending=False)
    combined = combined.astype(int)

    width = max(11, len(combined.index) * 1.6)
    fig, ax = plt.subplots(figsize=(width, 7))
    combined.plot(kind="bar", ax=ax)

    ax.set_title(f"{number} Частотное распределение: {column}")
    ax.set_xlabel("")
    ax.set_ylabel("Частота упоминаний")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
    ax.legend(title="Группа")
    add_bar_labels(ax)

    safe_col = column.replace(" ", "_").replace("/", "_")
    save_plot(f"{number}_{safe_col}_частоты.png")
    combined.to_csv(TABLES_DIR / f"{number}_{safe_col}_частоты.csv", encoding="utf-8-sig")


def plot_percent_stacked(column: str, number: str, title: str):
    """100% stacked bar: структура категории по платформам."""
    exploded = explode_category(df, column)
    tab = pd.crosstab(exploded["Платформа"], exploded["Значение"], normalize="index") * 100
    tab = tab.reindex(["TG", "TikTok"]).dropna(how="all")

    fig, ax = plt.subplots(figsize=(12, 6))
    tab.plot(kind="bar", stacked=True, ax=ax)
    ax.set_title(f"{number} {title}")
    ax.set_xlabel("")
    ax.set_ylabel("Доля внутри платформы, %")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
    ax.legend(title=column, bbox_to_anchor=(1.02, 1), loc="upper left")
    save_plot(f"{number}_{column.replace(' ', '_')}_проценты_по_платформам.png")
    tab.to_csv(TABLES_DIR / f"{number}_{column.replace(' ', '_')}_проценты_по_платформам.csv", encoding="utf-8-sig")


def plot_heatmap(row_col: str, col_col: str, number: str, title: str):
    """Тепловая карта связи двух многозначных категорий."""
    rows = []
    for _, row in df.iterrows():
        row_values = split_multi_values(row.get(row_col))
        col_values = split_multi_values(row.get(col_col))
        for rv in row_values:
            for cv in col_values:
                rows.append({row_col: rv, col_col: cv})

    temp = pd.DataFrame(rows)
    if temp.empty:
        return

    tab = pd.crosstab(temp[row_col], temp[col_col])

    fig, ax = plt.subplots(figsize=(12, 7))
    sns.heatmap(tab, annot=True, fmt="d", cmap="Blues", ax=ax)
    ax.set_title(f"{number} {title}")
    ax.set_xlabel(col_col)
    ax.set_ylabel(row_col)
    save_plot(f"{number}_тепловая_карта_{row_col}_x_{col_col}.png")
    tab.to_csv(TABLES_DIR / f"{number}_тепловая_карта_{row_col}_x_{col_col}.csv", encoding="utf-8-sig")


# =========================
# 3. ОСНОВНЫЕ ГРАФИКИ 01–07
# =========================
plot_platform_distribution()
plot_frequency_by_platform("Тип успеха", "02")
plot_frequency_by_platform("Эмоция", "03")
plot_frequency_by_platform("Триггеры", "04")
plot_frequency_by_platform("Аргументация", "05")
plot_frequency_by_platform("Призыв к действию", "06")
plot_frequency_by_platform("Образ успеха", "07")


# =========================
# 4. ДОПОЛНИТЕЛЬНЫЕ ГРАФИКИ С ПРОЦЕНТАМИ И СРАВНЕНИЯМИ
# =========================
# 08 — структура типа успеха по платформам
plot_percent_stacked("Тип успеха", "08", "Структура типа успеха по платформам (%)")

# 09 — структура эмоций по платформам
plot_percent_stacked("Эмоция", "09", "Структура эмоциональной окраски по платформам (%)")

# 10 — структура образа успеха по платформам
plot_percent_stacked("Образ успеха", "10", "Образ успеха по платформам (%)")

# 11 — тепловая карта: тип успеха × триггеры
plot_heatmap("Тип успеха", "Триггеры", "11", "Связь типа успеха и триггеров")

# 12 — тепловая карта: эмоция × триггеры
plot_heatmap("Эмоция", "Триггеры", "12", "Связь эмоций и триггеров")

# 13 — доля публикаций с триггерами по платформам
trigger_share = df.groupby("Платформа")["Есть триггер"].mean().reindex(["TG", "TikTok"]) * 100
fig, ax = plt.subplots(figsize=(8, 6))
trigger_share.plot(kind="bar", ax=ax)
ax.set_title("13 Доля публикаций с триггерами по платформам (%)")
ax.set_xlabel("")
ax.set_ylabel("Доля публикаций, %")
ax.set_ylim(0, 100)
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
add_bar_labels(ax, fmt="{:.1f}")
save_plot("13_доля_публикаций_с_триггерами.png")
trigger_share.to_csv(TABLES_DIR / "13_доля_публикаций_с_триггерами.csv", encoding="utf-8-sig")

# 14 — доля публикаций с CTA по платформам
cta_share = df.groupby("Платформа")["Есть CTA"].mean().reindex(["TG", "TikTok"]) * 100
fig, ax = plt.subplots(figsize=(8, 6))
cta_share.plot(kind="bar", ax=ax)
ax.set_title("14 Доля публикаций с призывом к действию по платформам (%)")
ax.set_xlabel("")
ax.set_ylabel("Доля публикаций, %")
ax.set_ylim(0, 100)
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
add_bar_labels(ax, fmt="{:.1f}")
save_plot("14_доля_CTA_по_платформам.png")
cta_share.to_csv(TABLES_DIR / "14_доля_CTA_по_платформам.csv", encoding="utf-8-sig")

# 15 — среднее число подписчиков по типу успеха
success_exploded = explode_category(df, "Тип успеха")
subs_by_success = success_exploded.groupby("Значение")["Подписчики"].mean().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(10, 6))
subs_by_success.plot(kind="bar", ax=ax)
ax.set_title("15 Среднее число подписчиков по типу успеха")
ax.set_xlabel("")
ax.set_ylabel("Среднее число подписчиков")
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
add_bar_labels(ax, fmt="{:.0f}")
save_plot("15_подписчики_по_типу_успеха.png")
subs_by_success.to_csv(TABLES_DIR / "15_подписчики_по_типу_успеха.csv", encoding="utf-8-sig")

# 16 — среднее число подписчиков по образу успеха
image_exploded = explode_category(df, "Образ успеха")
subs_by_image = image_exploded.groupby("Значение")["Подписчики"].mean().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(11, 6))
subs_by_image.plot(kind="bar", ax=ax)
ax.set_title("16 Среднее число подписчиков по образу успеха")
ax.set_xlabel("")
ax.set_ylabel("Среднее число подписчиков")
ax.set_xticklabels(ax.get_xticklabels(), rotation=35, ha="right")
add_bar_labels(ax, fmt="{:.0f}")
save_plot("16_подписчики_по_образу_успеха.png")
subs_by_image.to_csv(TABLES_DIR / "16_подписчики_по_образу_успеха.csv", encoding="utf-8-sig")

# 17 — подписчики и наличие триггеров: возможный парадокс популярности
subs_by_trigger = df.groupby("Есть триггер")["Подписчики"].mean()
subs_by_trigger.index = ["Нет триггера" if i == 0 else "Есть триггер" for i in subs_by_trigger.index]
fig, ax = plt.subplots(figsize=(8, 6))
subs_by_trigger.plot(kind="bar", ax=ax)
ax.set_title("17 Среднее число подписчиков: есть ли триггер")
ax.set_xlabel("")
ax.set_ylabel("Среднее число подписчиков")
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
add_bar_labels(ax, fmt="{:.0f}")
save_plot("17_подписчики_и_наличие_триггеров.png")
subs_by_trigger.to_csv(TABLES_DIR / "17_подписчики_и_наличие_триггеров.csv", encoding="utf-8-sig")

# 18 — средняя длина текста по платформам
text_len = df.groupby("Платформа")["Длина текста"].mean().reindex(["TG", "TikTok"])
fig, ax = plt.subplots(figsize=(8, 6))
text_len.plot(kind="bar", ax=ax)
ax.set_title("18 Средняя длина текста по платформам")
ax.set_xlabel("")
ax.set_ylabel("Средняя длина текста, символы")
ax.set_xticklabels(ax.get_xticklabels(), rotation=0)
add_bar_labels(ax, fmt="{:.0f}")
save_plot("18_средняя_длина_текста_по_платформам.png")
text_len.to_csv(TABLES_DIR / "18_средняя_длина_текста_по_платформам.csv", encoding="utf-8-sig")

# 19 — топ каналов по подписчикам и доля публикаций с триггерами
channel_stats = (
    df.groupby(["Платформа", "Название канала"], as_index=False)
    .agg(
        Подписчики=("Подписчики", "max"),
        Доля_триггеров=("Есть триггер", "mean"),
        Доля_CTA=("Есть CTA", "mean"),
        Средняя_длина_текста=("Длина текста", "mean"),
        Количество_публикаций=("Текст", "count"),
    )
)
channel_stats["Доля_триггеров"] *= 100
channel_stats["Доля_CTA"] *= 100
channel_stats = channel_stats.sort_values("Подписчики", ascending=False)
channel_stats.to_csv(TABLES_DIR / "19_статистика_по_каналам.csv", index=False, encoding="utf-8-sig")

plot_data = channel_stats.head(15).sort_values("Подписчики", ascending=True)
fig, ax = plt.subplots(figsize=(12, 8))
ax.barh(plot_data["Название канала"], plot_data["Подписчики"])
ax.set_title("19 Топ каналов по подписчикам")
ax.set_xlabel("Подписчики")
ax.set_ylabel("")
save_plot("19_топ_каналов_по_подписчикам.png")

# 20 — scatter: подписчики × доля триггеров
fig, ax = plt.subplots(figsize=(10, 7))
for platform, group in channel_stats.groupby("Платформа"):
    ax.scatter(group["Подписчики"], group["Доля_триггеров"], label=platform, s=80)

for _, row in channel_stats.iterrows():
    name = str(row["Название канала"])
    if len(name) > 22:
        name = name[:22] + "..."
    ax.annotate(name, (row["Подписчики"], row["Доля_триггеров"]), fontsize=7, alpha=0.8)

ax.set_title("20 Подписчики и доля публикаций с триггерами")
ax.set_xlabel("Количество подписчиков")
ax.set_ylabel("Доля публикаций с триггерами, %")
ax.legend(title="Платформа")
save_plot("20_подписчики_x_доля_триггеров.png")


# =========================
# 5. СВОДНЫЙ EXCEL-ОТЧЁТ
# =========================
summary_path = OUTPUT_DIR / "summary_tables.xlsx"
with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Исходные данные", index=False)

    # частоты смысловых единиц
    for i, column in enumerate(semantic_columns, start=2):
        exploded = explode_category(df, column)
        freq = exploded["Значение"].value_counts().rename_axis(column).reset_index(name="Частота")
        freq["Процент"] = freq["Частота"] / freq["Частота"].sum() * 100
        freq.to_excel(writer, sheet_name=f"{i:02d}_{column[:20]}", index=False)

    channel_stats.to_excel(writer, sheet_name="Каналы", index=False)

print("Готово!")
print(f"Графики сохранены в папке: {GRAPHS_DIR}")
print(f"Таблицы сохранены в папке: {TABLES_DIR}")
print(f"Сводный Excel-отчет: {summary_path}")
