import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
from datetime import datetime

# ==================== НАСТРОЙКИ ====================
# Порядок колонок в выгрузке плана:
OUTPUT_COLS = ['Артикул площадки', 'Артикул поставщика', 'Площадка', 'Ссылка',
               'Наименование', 'Статус', 'Рейтинг', 'Кол-во отзывов',
               'Приоритет', 'Сезонный', 'Рекомендация']

VALID_PLATFORMS = [
    'лемана про', 'лемана про мп', 'мегастрой',
    'максидом', 'петрович', 'все инструменты'
]

# Приоритет площадок (для приоритетов 3 и 4, последний ключ сортировки):
# чем меньше число, тем выше в списке при прочих равных. Не указанные — в конце.
PLATFORM_PRIORITY = {
    'лемана про': 0,
    'лемана про мп': 1,
}

# Ключевые слова по умолчанию. Если в репозитории лежит seasonal_keywords.txt /
# lamp_keywords.txt — списки берутся оттуда (по одному слову/фразе на строку,
# строки с # — комментарии).
DEFAULT_SEASONAL_KEYWORDS = [
    'светильник светодиодный', 'на солнечной батарее', 'на солнечных батареях',
    'солар', 'solar', 'прожектор', 'светильник настенный', 'сад', 'столб',
    'уличный', 'датчик', 'таймер', 'садовый', 'налобный', 'кемпинговый',
    'ручной', 'перчатка', 'перчатки', 'фонарь'
]

DEFAULT_LAMP_KEYWORDS = [
    'люстра', 'люстр', 'потолочный', 'бра', 'светильник',
    'подвесной', 'торшер', 'спот', 'ночник'
]

SEASONALITY_FILES = ['seasonality.xlsx',
                     'lm_cats_distribution_consistent_filtered_cats.xlsx']

# Файл несезонных слов: слово ; месяцы провала (1-12)
OFFSEASON_FILE = 'seasonal_offseason.txt'


def parse_offseason_lines(lines):
    """Список (слово, regex, {месяцы}), отсортированный от длинных фраз к коротким.

    Многословные фразы допускают вставки между словами: фраза «фонарь настенный»
    найдётся и в «фонарь-подсветка настенный». Левая граница слова сохраняется,
    чтобы 'сад' не цеплял 'фасад'.
    """
    items = []
    for line in lines:
        line = line.strip()
        if not line or line.startswith('#') or ';' not in line:
            continue
        parts = line.split(';')
        word = parts[0].strip().lower()
        months = set()
        for p in parts[1:]:
            p = p.strip()
            if p.isdigit() and 1 <= int(p) <= 12:
                months.add(int(p))
        if not (word and months):
            continue
        tokens = word.split()
        if len(tokens) > 1:
            # слова фразы по порядку; между ними до ~25 символов вставок
            pat = r'(?<!\w)' + r'\w*\b.{0,25}?\b'.join(re.escape(t) for t in tokens)
        else:
            pat = r'(?<!\w)' + re.escape(word)
        items.append((word, re.compile(pat), months))
    # длинные фразы первыми — конкретное побеждает общее
    items.sort(key=lambda x: -len(x[0]))
    return items


def load_offseason_file(path=OFFSEASON_FILE):
    if not os.path.exists(path):
        return [], f"Файл {path} не найден — сезонный фильтр не применяется."
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return parse_offseason_lines(f), None
    except Exception as e:
        return [], f"Не удалось прочитать {path}: {e}"


def is_offseason(name_lower, month, offseason_items):
    """True, если для названия сработало несезонное слово в данном месяце.

    Идём от длинных фраз к коротким — первое совпавшее слово решает.
    """
    if not month:
        return False
    for word, pattern, months in offseason_items:
        if pattern.search(name_lower):
            return month in months
    return False


MONTH_NAMES = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
               'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']


def make_patterns(keywords):
    """Слева — граница слова: 'сад' найдёт 'сад'/'садовый', но не 'фасадный'."""
    return [re.compile(r'(?<!\w)' + re.escape(kw.lower())) for kw in keywords]


def matches_any(text, patterns):
    return any(p.search(text) for p in patterns)


def load_keywords(path, default):
    """Список ключевых слов из файла; если файла нет — встроенный список."""
    if not os.path.exists(path):
        return default, None
    try:
        with open(path, 'r', encoding='utf-8') as f:
            kws = [line.strip().lower() for line in f
                   if line.strip() and not line.strip().startswith('#')]
        if not kws:
            return default, f"{path} пуст — использован встроенный список."
        return kws, None
    except Exception as e:
        return default, f"Не удалось прочитать {path}: {e}. Использован встроенный список."


# ==================== ЧЁРНЫЙ СПИСОК ====================
def normalize_sku(value):
    s = str(value).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s


def blacklist_from_lines(lines):
    bl = set()
    for line in lines:
        sku = normalize_sku(line)
        if sku and not sku.startswith('#'):
            bl.add(sku)
            bl.add(sku.replace(' ', ''))
    return bl


def load_blacklist_file(path='blacklist.txt'):
    if not os.path.exists(path):
        return set(), f"Файл {path} не найден — чёрный список пуст."
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return blacklist_from_lines(f), None
    except Exception as e:
        return set(), f"Не удалось прочитать {path}: {e}"


# ==================== СЕЗОННОСТЬ (только справочно) ====================
@st.cache_data(show_spinner=False)
def parse_seasonality(file_bytes):
    """category -> {1..12: доля месяца}. На отбор НЕ влияет, выводится как подсказка."""
    sdf = pd.read_excel(io.BytesIO(file_bytes))
    sdf.columns = [str(c).strip() for c in sdf.columns]
    if 'category' not in sdf.columns:
        raise ValueError("нет колонки 'category'")
    month_cols = {}
    for c in sdf.columns:
        try:
            m = int(float(c))
        except ValueError:
            continue
        if 1 <= m <= 12:
            month_cols[m] = c
    if len(month_cols) != 12:
        raise ValueError(f"найдено {len(month_cols)} месячных колонок вместо 12")
    return {str(r['category']).strip(): {m: float(r[col]) for m, col in month_cols.items()}
            for _, r in sdf.iterrows()}


def load_seasonality_file():
    for path in SEASONALITY_FILES:
        if os.path.exists(path):
            try:
                with open(path, 'rb') as f:
                    return parse_seasonality(f.read()), None
            except Exception as e:
                return {}, f"Файл сезонности {path}: {e}"
    return {}, None


# ==================== ПАРСИНГ ОТЧЁТА (с кешем) ====================
@st.cache_data(show_spinner=False)
def parse_csv(file_bytes):
    last_err = None
    for enc in ('utf-8', 'cp1251'):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), sep=None,
                               engine='python', encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
    raise ValueError(f"Не удалось определить кодировку CSV (UTF-8/CP1251): {last_err}")


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes):
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def parse_excel(file_bytes, sheet_name):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)


# ==================== ОБРАБОТКА ====================
def clean_numeric(val):
    if pd.isna(val):
        return np.nan
    val_str = str(val).replace('\xa0', '').replace(' ', '').replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return np.nan


def check_status(status):
    """False = закрыт к заказам, строка отсеивается."""
    if pd.isna(status):
        return True
    s = str(status).strip().lower()
    if 'закрыт к' in s:
        return False
    return True


def is_sale_status(status):
    """Распродажа / вывод товара (даже если открыт к заказам)."""
    if pd.isna(status):
        return False
    s = str(status).lower()
    return ('распродаж' in s) or ('вывод' in s)


def get_priority(row, target_rating):
    rating = row['Рейтинг_число']
    prev_rating = row['Предыдущий_рейтинг_число']
    reviews = row['Отзывы_число']

    # 1: критично низкий рейтинг
    if pd.notna(rating) and rating <= 3.9:
        return 1
    # 2: спад относительно предыдущего рейтинга
    if pd.notna(rating) and rating >= 4.0 and pd.notna(prev_rating) and rating < prev_rating:
        return 2
    # 3: ниже проходного и отзывов <= 4
    if pd.notna(rating) and rating <= target_rating and (pd.isna(reviews) or reviews <= 4):
        return 3
    # 4: высокий/неизвестный рейтинг и отзывов <= 2 — риск обвала
    if (pd.isna(rating) or rating > 4.3) and (pd.isna(reviews) or reviews <= 2):
        return 4
    return 99


def get_recommendation(row):
    rating = row['Рейтинг_число']
    reviews = row['Отзывы_число']
    if pd.notna(rating) and rating >= 4.0 and pd.notna(reviews) and reviews == 1:
        return "⚠️ Риск падения: написать 1-2 отзыва"
    if pd.notna(rating) and rating < 4.0 and pd.notna(reviews) and reviews >= 10:
        return "📅 Долгосрочная работа: 3-5 отзывов постепенно"
    return "✅ Стандартная проработка"


PRIORITY_EMOJI = {1: "🔴 1", 2: "🟠 2", 3: "🟡 3", 4: "🔵 4"}


def find_link_column(columns):
    """Ищем в отчёте колонку с прямыми ссылками ('Ссылка', 'URL' и т.п.)."""
    for col in columns:
        c = str(col).strip().lower()
        if 'ссылк' in c or c == 'url' or 'url' in c.split():
            return col
    return None


def find_stock_column(columns):
    """Ищем колонку наличия/остатка."""
    for col in columns:
        c = str(col).strip().lower()
        if 'наличи' in c or 'остат' in c or c == 'сток' or 'склад' in c:
            return col
    return None


def find_date_column(columns):
    """Ищем колонку с датой парсинга."""
    for col in columns:
        c = str(col).strip().lower()
        if 'дата парсинга' in c or 'дата' in c or 'парсинг' in c:
            return col
    return None


def detect_report_month(series):
    """Месяц отчёта = самая частая дата в колонке (формат ДД.ММ.ГГГГ и пр.)."""
    parsed = pd.to_datetime(series, dayfirst=True, errors='coerce')
    parsed = parsed.dropna()
    if parsed.empty:
        return None
    return int(parsed.dt.month.mode().iloc[0])


# Текстовые значения, означающие отсутствие товара
OUT_OF_STOCK_WORDS = ['нет в наличии', 'не в наличии', 'отсутств', 'нет',
                      'под заказ', 'распродан', 'закончил', 'out of stock', '0']


def is_in_stock(value):
    """True, если товар в наличии.

    Формат отчёта: колонка «Наличие» со значениями «В наличии» / «Нет в наличии».
    Дополнительно поддержаны числовые остатки и прочие текстовые формулировки.
    """
    if pd.isna(value):
        return False
    s = str(value).strip().lower()
    if s == '':
        return False
    # точные значения из отчёта
    if s == 'в наличии':
        return True
    if s == 'нет в наличии':
        return False
    # число: 0 -> нет, >0 -> есть
    num = clean_numeric(s)
    if pd.notna(num):
        return num > 0
    # прочий текст: ищем слова об отсутствии
    if any(w in s for w in OUT_OF_STOCK_WORDS):
        return False
    return True


# ==================== ИНТЕРФЕЙС ====================
st.set_page_config(page_title="Аналитика Рейтингов 4.5+", layout="wide")

# ---------- Шапка: два изображения рядом, по центру ----------
_, img_left, img_right, _ = st.columns([2, 1, 1, 2])
if os.path.exists('logo.png'):
    with img_left:
        st.image('logo.png', use_container_width=True)
for photo in ('второе_фото.png', 'photo2.png'):
    if os.path.exists(photo):
        with img_right:
            st.image(photo, use_container_width=True)
        break

st.markdown(
    "<h1 style='text-align: center;'>✨ Автоматизация отбора артикулов "
    "для работы с рейтингом</h1>",
    unsafe_allow_html=True
)

# ---------- Сайдбар ----------
st.sidebar.header("⚙️ Настройки")
target_rating = st.sidebar.number_input(
    "Проходной рейтинг (приоритет 3)", 4.0, 5.0, 4.5, 0.1,
    help="Повышайте по мере исправления ситуации.")
max_rows = int(st.sidebar.number_input(
    "Лимит объёма: строк «артикул × площадка»", 10, 1000, 150, 10,
    help="Если лимит срезает артикул посередине, он добирается целиком."))

st.sidebar.subheader("🛍️ Распродажа")
show_sale = st.sidebar.checkbox(
    "Показать товары на распродаже", value=False,
    help="Справочная таблица под основным планом: распродажа/вывод с низким "
         "рейтингом. В основной план такие товары попадают только с ≤2 отзывами.")
sale_view_reviews = 5
if show_sale:
    sale_view_reviews = int(st.sidebar.number_input(
        "Отзывов не более", 1, 20, 5, 1))

bl_upload = st.sidebar.file_uploader("Обновить чёрный список (txt)", type=['txt'])
kw_upload = st.sidebar.file_uploader("Обновить сезонные слова (txt)", type=['txt'])
off_upload = st.sidebar.file_uploader("Обновить несезонные слова по месяцам (txt)",
                                      type=['txt'])

# ---------- Чёрный список ----------
if bl_upload is not None:
    try:
        lines = bl_upload.getvalue().decode('utf-8', errors='replace').splitlines()
        BLACKLIST = blacklist_from_lines(lines)
        st.sidebar.success(f"Чёрный список: {len(BLACKLIST) // 2} артикулов (из файла)")
    except Exception as e:
        BLACKLIST = set()
        st.sidebar.error(f"Не удалось прочитать список: {e}")
else:
    BLACKLIST, bl_warning = load_blacklist_file()
    if bl_warning:
        st.sidebar.warning(bl_warning)

# ---------- Ключевые слова ----------
if kw_upload is not None:
    lines = kw_upload.getvalue().decode('utf-8', errors='replace').splitlines()
    SEASONAL_KEYWORDS = [l.strip().lower() for l in lines
                         if l.strip() and not l.strip().startswith('#')]
    st.sidebar.success(f"Сезонные слова: {len(SEASONAL_KEYWORDS)} (из файла)")
else:
    SEASONAL_KEYWORDS, kw_warning = load_keywords('seasonal_keywords.txt',
                                                  DEFAULT_SEASONAL_KEYWORDS)
    if kw_warning:
        st.sidebar.warning(kw_warning)

LAMP_KEYWORDS, lamp_warning = load_keywords('lamp_keywords.txt', DEFAULT_LAMP_KEYWORDS)
if lamp_warning:
    st.sidebar.warning(lamp_warning)

SEASONAL_PATTERNS = make_patterns(SEASONAL_KEYWORDS)
LAMP_PATTERNS = make_patterns(LAMP_KEYWORDS)

# ---------- Несезонные слова по месяцам (новый фильтр) ----------
if off_upload is not None:
    lines = off_upload.getvalue().decode('utf-8', errors='replace').splitlines()
    OFFSEASON_ITEMS = parse_offseason_lines(lines)
    st.sidebar.success(f"Несезонные слова: {len(OFFSEASON_ITEMS)} (из файла)")
else:
    OFFSEASON_ITEMS, off_warning = load_offseason_file()
    if off_warning:
        st.sidebar.warning(off_warning)

apply_offseason = st.sidebar.checkbox(
    "Учитывать сезонность по месяцам", value=True,
    help="Несезонные товары: приоритеты 3-4 убираем; приоритет 1 оставляем при "
         "<5 отзывах; приоритет 2 — при <10 отзывах.")

st.caption(f"Цель: {target_rating}+ на всех площадках | "
           f"Чёрный список: {len(BLACKLIST) // 2 if BLACKLIST else 0} артикулов | "
           f"Несезонных правил: {len(OFFSEASON_ITEMS)}")

uploaded_file = st.file_uploader("📁 Загрузите еженедельный отчет (CSV или Excel)",
                                 type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.getvalue()
        if uploaded_file.name.endswith('.csv'):
            df = parse_csv(file_bytes)
        else:
            sheet_names = get_sheet_names(file_bytes)
            default_index = 0
            for i, name in enumerate(sheet_names):
                if 'текущ' in name.lower():
                    default_index = i
                    break
            selected_sheet = st.selectbox("📂 Выберите лист:", sheet_names,
                                          index=default_index)
            df = parse_excel(file_bytes, selected_sheet)

        st.success("✅ Файл успешно прочитан!")
        df.columns = df.columns.astype(str).str.strip()

        required_cols = ['Артикул поставщика', 'Статус', 'Площадка', 'Рейтинг',
                         'Кол-во отзывов', 'Предыдущий рейтинг', 'Наименование']
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            st.error(f"❌ Не найдены колонки: {', '.join(missing_cols)}")
        else:
            with st.spinner("🪄 Анализируем данные..."):
                df['Рейтинг_число'] = df['Рейтинг'].apply(clean_numeric)
                df['Отзывы_число'] = df['Кол-во отзывов'].apply(clean_numeric)
                df['Предыдущий_рейтинг_число'] = df['Предыдущий рейтинг'].apply(clean_numeric)
                df['СКУ_норм'] = df['Артикул поставщика'].apply(normalize_sku)

                # --- Воронка фильтрации ---
                n_total = len(df)
                df_f = df[df['Статус'].apply(check_status)].copy()
                n_status = len(df_f)

                df_f = df_f[df_f['Площадка'].astype(str).str.strip()
                            .str.lower().isin(VALID_PLATFORMS)]
                n_platform = len(df_f)

                in_blacklist = (
                    df_f['СКУ_норм'].isin(BLACKLIST) |
                    df_f['СКУ_норм'].str.replace(' ', '', regex=False).isin(BLACKLIST)
                )
                n_blacklisted = int(in_blacklist.sum())
                df_f = df_f[~in_blacklist]

                # --- Фильтр наличия: товары не в наличии в работу не берём ---
                stock_col = find_stock_column(df_f.columns)
                n_out_of_stock = 0
                if stock_col:
                    in_stock_mask = df_f[stock_col].apply(is_in_stock)
                    n_out_of_stock = int((~in_stock_mask).sum())
                    df_f = df_f[in_stock_mask]

                # --- Приоритеты ---
                df_f['Приоритет'] = df_f.apply(get_priority, axis=1,
                                               target_rating=target_rating)
                df_f['Распродажа'] = df_f['Статус'].apply(is_sale_status)
                # копия всех распродажных строк — для справочной таблицы
                df_sale_all = df_f[df_f['Распродажа']].copy()

                # Распродажа/вывод: берём в работу только приоритет 1 с <3 отзывами
                sale_drop = df_f['Распродажа'] & ~(
                    (df_f['Приоритет'] == 1) &
                    (df_f['Отзывы_число'].fillna(0) <= 2)
                )
                # считаем отсев только среди потенциально проблемных строк
                n_sale_dropped = int((sale_drop & (df_f['Приоритет'] <= 4)).sum())
                df_f = df_f[~sale_drop]

                # --- Сезонный фильтр по месяцу даты парсинга ---
                report_month = None
                date_col = find_date_column(df.columns)
                if date_col is not None:
                    report_month = detect_report_month(df[date_col])

                n_offseason_dropped = 0
                if apply_offseason and OFFSEASON_ITEMS and report_month:
                    names_low = df_f['Наименование'].astype(str).str.lower()
                    off_mask = names_low.apply(
                        lambda n: is_offseason(n, report_month, OFFSEASON_ITEMS))
                    rev = df_f['Отзывы_число'].fillna(0)
                    # несезонные оставляем по правилам приоритетов:
                    keep_offseason = (
                        ((df_f['Приоритет'] == 1) & (rev < 5)) |
                        ((df_f['Приоритет'] == 2) & (rev < 10))
                    )
                    drop_offseason = off_mask & ~keep_offseason & (df_f['Приоритет'] <= 4)
                    n_offseason_dropped = int(drop_offseason.sum())
                    df_f = df_f[~drop_offseason]

                df_problems = df_f[df_f['Приоритет'] <= 4].copy()
                n_problems = len(df_problems)

                with st.expander("🔎 Воронка фильтрации"):
                    if report_month:
                        st.write(f"Месяц отчёта (по дате парсинга): "
                                 f"**{MONTH_NAMES[report_month - 1]}**")
                    elif apply_offseason:
                        st.info("Колонка с датой парсинга не найдена или дата не "
                                "распознана — сезонный фильтр не применялся.")
                    st.write(f"Строк в файле: **{n_total}**")
                    st.write(f"Отсеяно по статусу «закрыт к заказам»: **{n_total - n_status}**")
                    st.write(f"Отсеяно по площадке: **{n_status - n_platform}**")
                    st.write(f"Исключено чёрным списком: **{n_blacklisted}**")
                    if stock_col:
                        st.write(f"Не в наличии (колонка «{stock_col}»): "
                                 f"**{n_out_of_stock}**")
                    else:
                        st.info("Колонка наличия не найдена — фильтр по наличию "
                                "не применялся.")
                    st.write(f"Распродажа/вывод, работа нерациональна: **{n_sale_dropped}**")
                    st.write(f"Отложено как несезонные: **{n_offseason_dropped}**")
                    st.write(f"Не требуют работы (приоритет 99): "
                             f"**{len(df_f) - n_problems}**")
                    if BLACKLIST and n_blacklisted == 0:
                        st.warning("Чёрный список загружен, но ни одна строка не исключена — "
                                   "проверьте формат артикулов в списке и в отчёте.")

                # ==================== РАСПРОДАЖА (справочно) ====================
                if show_sale:
                    st.markdown("---")
                    st.subheader("🛍️ Товары на распродаже")
                    sale_view = df_sale_all[
                        (df_sale_all['Рейтинг_число'] <= 3.9) &
                        (df_sale_all['Отзывы_число'].fillna(0) <= sale_view_reviews)
                    ].copy()
                    if sale_view.empty:
                        st.info("Распродажных товаров с рейтингом ≤3.9 и отзывами "
                                f"≤{sale_view_reviews} не найдено.")
                    else:
                        s_link = find_link_column(sale_view.columns)
                        sale_view['Ссылка'] = (
                            sale_view[s_link].astype(str).str.strip()
                            .replace({'nan': '', 'None': ''}) if s_link else '')
                        sale_view['В основном плане'] = (
                            (sale_view['Приоритет'] == 1) &
                            (sale_view['Отзывы_число'].fillna(0) <= 2))
                        sale_cols = [c for c in
                                     ['Артикул площадки', 'Артикул поставщика',
                                      'Площадка', 'Ссылка', 'Наименование', 'Статус',
                                      'Рейтинг', 'Кол-во отзывов', 'В основном плане']
                                     if c in sale_view.columns]
                        st.caption(f"Распродажа/вывод, рейтинг ≤3.9, отзывов "
                                   f"≤{sale_view_reviews}. Справочно — на основной "
                                   f"план не влияет.")
                        st.dataframe(
                            sale_view[sale_cols], use_container_width=True,
                            column_config={
                                'Ссылка': st.column_config.LinkColumn(
                                    'Ссылка', display_text='Открыть 🔗')
                            })

                if df_problems.empty:
                    st.warning("⚠️ Проблемных артикулов не обнаружено.")
                else:
                    # Финальный приоритет артикула = минимальный по всем площадкам
                    df_problems['Приоритет'] = (df_problems
                                                .groupby('СКУ_норм')['Приоритет']
                                                .transform('min'))

                    names_lower = df_problems['Наименование'].astype(str).str.lower()
                    df_problems['Сезонный'] = names_lower.apply(
                        lambda n: matches_any(n, SEASONAL_PATTERNS))
                    is_lamp = names_lower.apply(lambda n: matches_any(n, LAMP_PATTERNS))

                    # Внутри приоритета 4: сезонные слова -> люстры -> остальное
                    df_problems['Сортировка_4'] = np.select(
                        [
                            (df_problems['Приоритет'] == 4) & df_problems['Сезонный'],
                            (df_problems['Приоритет'] == 4) & is_lamp,
                            (df_problems['Приоритет'] == 4),
                        ],
                        [1, 2, 3],
                        default=0
                    )

                    # Значимость площадки — последний ключ, только для приоритетов 3-4.
                    # Для приоритетов 1-2 ключ = 0 у всех (не влияет).
                    plat_norm = df_problems['Площадка'].astype(str).str.strip().str.lower()
                    plat_rank = plat_norm.map(PLATFORM_PRIORITY).fillna(99)
                    df_problems['Сортировка_площадка'] = np.where(
                        df_problems['Приоритет'] >= 3, plat_rank, 0)

                    df_problems.sort_values(
                        by=['Приоритет', 'Сортировка_4', 'Сортировка_площадка',
                            'Рейтинг_число'],
                        ascending=[True, True, True, True],
                        inplace=True
                    )

                    # --- Лимит в строках «артикул × площадка», артикул целиком ---
                    sku_order = df_problems['СКУ_норм'].drop_duplicates().tolist()
                    sku_counts = df_problems['СКУ_норм'].value_counts()
                    selected_skus, total_rows = [], 0
                    for sku in sku_order:
                        if total_rows >= max_rows:
                            break
                        selected_skus.append(sku)
                        total_rows += int(sku_counts[sku])

                    final_df = df_problems[df_problems['СКУ_норм']
                                           .isin(selected_skus)].copy()
                    final_df['Рекомендация'] = final_df.apply(get_recommendation, axis=1)

                    # --- Ссылка на карточку: только из отчёта ---
                    report_link_col = find_link_column(final_df.columns)
                    if report_link_col:
                        final_df['Ссылка'] = (final_df[report_link_col]
                                              .astype(str).str.strip()
                                              .replace({'nan': '', 'None': ''}))
                    else:
                        final_df['Ссылка'] = ''
                        st.warning("В отчёте не найдена колонка со ссылками "
                                   "(«Ссылка», «URL» и т.п.) — колонка останется пустой.")

                    # --- Артикул площадки ---
                    if 'Артикул площадки' not in final_df.columns:
                        final_df['Артикул площадки'] = ''
                        st.warning("В отчёте не найдена колонка «Артикул площадки» — "
                                   "колонка останется пустой.")

                    # ==================== ДАШБОРД ====================
                    st.markdown("---")
                    st.subheader("💫 Дашборд")

                    col1, col2, col3, col4 = st.columns(4)
                    sku_level = final_df.drop_duplicates('СКУ_норм')
                    with col1:
                        st.metric("🔴 Приоритет 1 (≤3.9)",
                                  len(sku_level[sku_level['Приоритет'] == 1]))
                    with col2:
                        st.metric("🟠 Приоритет 2 (спад)",
                                  len(sku_level[sku_level['Приоритет'] == 2]))
                    with col3:
                        st.metric(f"🟡 Приоритет 3 (≤{target_rating}, отзывов ≤4)",
                                  len(sku_level[sku_level['Приоритет'] == 3]))
                    with col4:
                        st.metric("🔵 Приоритет 4 (риск, отзывов ≤2)",
                                  len(sku_level[sku_level['Приоритет'] == 4]))

                    st.write(f"**Объём плана:** {len(final_df)} строк "
                             f"«артикул × площадка» (лимит {max_rows}) | "
                             f"**Уникальных артикулов:** {len(selected_skus)} | "
                             f"**Исключено чёрным списком:** {n_blacklisted} строк")

                    # ==================== ТАБЛИЦА ====================
                    st.markdown("---")
                    st.subheader("📝 План работы")

                    output_df = final_df[OUTPUT_COLS].copy()

                    screen_df = output_df.copy()
                    screen_df['Приоритет'] = screen_df['Приоритет'].map(
                        lambda p: PRIORITY_EMOJI.get(p, str(p)))

                    st.dataframe(
                        screen_df, use_container_width=True, height=600,
                        column_config={
                            'Ссылка': st.column_config.LinkColumn(
                                'Ссылка', display_text='Открыть 🔗')
                        })

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        output_df.to_excel(writer, index=False,
                                           sheet_name='План Отзывов')
                        ws = writer.sheets['План Отзывов']
                        link_col_idx = output_df.columns.get_loc('Ссылка') + 1
                        for row_i, url in enumerate(output_df['Ссылка'], start=2):
                            if url:
                                cell = ws.cell(row=row_i, column=link_col_idx)
                                cell.value = 'Открыть'
                                cell.hyperlink = url
                                cell.style = 'Hyperlink'
                    st.download_button(
                        label="📥 Скачать план в Excel",
                        data=buffer.getvalue(),
                        file_name="План_работ_по_рейтингам.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
