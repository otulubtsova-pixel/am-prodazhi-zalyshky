"""
Веб-версия generate_report.py на Streamlit.

Запуск локально:
    pip3 install -r requirements.txt
    streamlit run app.py

Публикация для других пользователей: streamlit.app (Streamlit Community
Cloud) - запушить этот репозиторий на GitHub и подключить его на
share.streamlit.io, указав app.py как главный файл.
"""

import io
from datetime import datetime

import streamlit as st

from generate_report import build_report

st.set_page_config(page_title="Звід АМ + продажі + залишки", page_icon="📦", layout="centered")

st.title("Звід АМ + продажі + залишки")
st.write(
    "Завантажте файл контрагента (АМ+продажі+залишки) та графік поставок — "
    "отримаєте зведений файл з полями «В дорозі», кандидатами в новинки "
    "та підсвіченою колонкою «НОВА АМ» по кожному магазину."
)

col1, col2 = st.columns(2)
with col1:
    src_file = st.file_uploader(
        "Файл контрагента (.xlsx)",
        type=["xlsx"],
        help='Наприклад "АМ+продажі+залишки_Міленіум_контрагент.xlsx" — з листами "АМ" і "Продажі 2025".',
    )
with col2:
    delivery_file = st.file_uploader(
        "Графік поставок (.xls)",
        type=["xls"],
        help='Наприклад "Графік поставок 04,09,2026.xls" — старий формат Excel (.xls), листи Chicco/Kids2/Offspring/Kendamil та інші.',
    )

if st.button("Сформувати звіт", type="primary", disabled=not (src_file and delivery_file)):
    try:
        with st.spinner("Обробляю файли..."):
            src_bytes = io.BytesIO(src_file.getvalue())
            delivery_bytes = delivery_file.getvalue()
            out_wb, stats = build_report(src_bytes, delivery_bytes)

            buf = io.BytesIO()
            out_wb.save(buf)
            buf.seek(0)
    except KeyError as e:
        st.error(f"У файлі контрагента не знайдено очікуваний лист: {e}. "
                 f"Перевірте, що є листи \"АМ\" і \"Продажі 2025\".")
    except Exception as e:
        st.error(f"Не вдалося сформувати звіт: {e}")
    else:
        st.success("Готово!")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Товарів", stats["n_products"])
        c2.metric("Магазинів", stats["n_stores"])
        c3.metric("Рядків товарів", stats["n_rows"])
        c4.metric("Кандидатів у новинки", stats["n_novelty"])

        out_name = f"Звід_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
        st.download_button(
            "Завантажити звіт (.xlsx)",
            data=buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
with st.expander("Що робить звіт (коротко)"):
    st.markdown(
        """
- Один рядок на товар, магазини — блоками колонок (як у вихідному файлі).
- Періоди продажів **«11 Лис. - 01 Січ.»** та **«02 Лют. - 10 Жов.»** — сума
  кількості проданого за наявні в файлі місяці цього періоду.
- **«АМ»** і **«Залишок»** — з листа «АМ» по кожному магазину.
- **«НОВА АМ»** — поки порожнє поле (світло-зелена заливка), заповнюється вручну.
- **«В дорозі»** — сума кількості з графіка поставок за артикулом (усі 4 листи:
  Chicco/Kids2/Offspring/Kendamil). Порожньо = артикула немає в графіку;
  0 = є, але зараз нічого не їде.
- Кандидати в **новинки** — окремі рядки внизу (лише «Наименование товара»,
  «Артикул», «В дорозі»): шукаються тільки на листі Chicco (єдиний бренд,
  де підтверджено, що порожній «Статус артикула» означає новинку), і лише
  якщо бренд Chicco взагалі є в асортименті цього контрагента.
"""
    )
