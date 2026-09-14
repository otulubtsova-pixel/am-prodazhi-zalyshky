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
from pathlib import Path

import streamlit as st

from generate_report import build_epicentr_combined_report, build_report

st.set_page_config(page_title="Звід АМ + продажі + залишки", page_icon="📦", layout="centered")

FILE_TYPES = ["xlsx", "xls"]

tab_antoshka, tab_epicentr = st.tabs(["Антошка", "Епіцентр"])

with tab_antoshka:
    st.title("Звід АМ + продажі + залишки")
    st.write(
        "Завантажте файл контрагента (АМ+продажі+залишки), графік поставок і прайс-лист — "
        "отримаєте зведений файл з полями «В дорозі», «Статус артикула», «Склад», кандидатами "
        "в новинки та підсвіченою колонкою «НОВА АМ» по кожному магазину."
    )

    col1, col2 = st.columns(2)
    with col1:
        src_file = st.file_uploader(
            "Файл контрагента (.xlsx або .xls)",
            type=FILE_TYPES,
            help='Наприклад "АМ+продажі+залишки_Міленіум_контрагент.xlsx" — з листами "АМ" і "Продажі 2025".',
            key="antoshka_src",
        )
    with col2:
        delivery_file = st.file_uploader(
            "Графік поставок (.xls або .xlsx)",
            type=FILE_TYPES,
            help='Наприклад "Графік поставок 04,09,2026.xls" — листи Chicco/Kids2/Offspring/Kendamil та інші.',
            key="antoshka_delivery",
        )

    price_file = st.file_uploader(
        "Прайс-лист з 1С УТП (.xlsx або .xls)",
        type=FILE_TYPES,
        help='Наприклад "Прайс-лист 11.09.2026.xlsx" — звіт «Прайс-лист (СКД)». Додає «Статус артикула» '
             'і «Склад» (загальний залишок), співставляючи за «Артикул».',
        key="antoshka_price",
    )

    abc_files = st.file_uploader(
        "Файли АВС-аналізу продажів з 1С УТП (можна декілька)",
        type=FILE_TYPES,
        accept_multiple_files=True,
        help='Наприклад "ABC Chicco.xlsx" — звіт «АВС-аналіз продажів (за номенклатурою)», '
             'по одному файлу на кожну групу товарів. Додає поле «Категорія» («АВС-клас») — '
             'співставляючи за «Артикул», а якщо не співпало — за точною «Назва».',
        key="antoshka_abc",
    )

    if st.button(
        "Сформувати звід",
        type="primary",
        disabled=not (src_file and delivery_file and price_file and abc_files),
        key="antoshka_btn",
    ):
        try:
            with st.spinner("Обробляю файли..."):
                out_wb, stats = build_report(src_file, delivery_file, price_file, abc_files)

                buf = io.BytesIO()
                out_wb.save(buf)
                buf.seek(0)
        except KeyError as e:
            st.error(f"У файлі контрагента не знайдено очікуваний лист: {e}. "
                     f"Перевірте, що є листи \"АМ\" і \"Продажі 2025\".")
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Не вдалося сформувати звід: {e}")
        else:
            st.success("Готово!")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Товарів", stats["n_products"])
            c2.metric("Магазинів", stats["n_stores"])
            c3.metric("Рядків товарів", stats["n_rows"])
            c4.metric("Кандидатів у новинки", stats["n_novelty"])

            out_name = f"Звід_{Path(src_file.name).stem}.xlsx"
            st.download_button(
                "Завантажити звід (.xlsx)",
                data=buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="antoshka_download",
            )

    st.divider()
    with st.expander("Що робить звід (коротко)"):
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
- **«Статус артикула»** і **«Склад»** — підтягуються з прайс-листа за
  збігом «Артикул» («Склад» = загальний залишок «Итого» з прайсу). Якщо
  для конкретного товару збігу немає — поля порожні.
- **«Категорія»** — підтягується з файлів АВС-аналізу («АВС-клас») за
  збігом «Артикул», а якщо не співпало — за точною «Назва» (без жодних
  відхилень). Можна завантажити декілька файлів АВС-аналізу відразу —
  результати об'єднуються.
- Кандидати в **новинки** — окремі рядки внизу (лише «Наименование товара»,
  «Артикул», «В дорозі»): шукаються тільки на листі Chicco (єдиний бренд,
  де підтверджено, що порожній «Статус артикула» означає новинку), і лише
  якщо бренд Chicco взагалі є в асортименті цього контрагента.
"""
        )

with tab_epicentr:
    st.title("Звід Епіцентр")
    st.write(
        "Завантажте по два звіти «Звіт про продажу товарів» (сезон / не сезон) для кожного "
        "контрагента — Мілленіум Трейд (МТ) і БебіШоп (БШ), файл-довідник попереднього періоду "
        "(для «АМ»), прайс-лист (для «Артикул»/«Бренд»/«Статус артикула»/«Склад»), графік "
        "поставок (для «В дорозі») та файли АВС-аналізу (для «Категорія»). Отримаєте один файл "
        "з двома листами («МТ» і «БШ»)."
    )

    st.subheader("Мілленіум Трейд (МТ)")
    col1, col2 = st.columns(2)
    with col1:
        mt_season_file = st.file_uploader(
            "Файл за сезон (.xls або .xlsx)",
            type=FILE_TYPES,
            help='Наприклад "Мілленіум трейд 01.01.2026-31.01.2026.xls".',
            key="epicentr_mt_season",
        )
    with col2:
        mt_offseason_file = st.file_uploader(
            "Файл за не сезон (.xls або .xlsx)",
            type=FILE_TYPES,
            help='Наприклад "Мілленіум трейд 01.08.2026-31.08.2026.xls".',
            key="epicentr_mt_offseason",
        )

    st.subheader("БебіШоп (БШ)")
    col3, col4 = st.columns(2)
    with col3:
        bsh_season_file = st.file_uploader(
            "Файл за сезон (.xls або .xlsx)",
            type=FILE_TYPES,
            help='Наприклад "БЕбішоп 01.01.2026-31.01.2026.xls".',
            key="epicentr_bsh_season",
        )
    with col4:
        bsh_offseason_file = st.file_uploader(
            "Файл за не сезон (.xls або .xlsx)",
            type=FILE_TYPES,
            help='Наприклад "БЕбішоп 01.08.2026-31.08.2026.xls".',
            key="epicentr_bsh_offseason",
        )

    st.subheader("Додаткові файли")
    col5, col6, col7 = st.columns(3)
    with col5:
        reference_file = st.file_uploader(
            "Довідник попереднього періоду",
            type=FILE_TYPES,
            help='Наприклад "Епік_матриця_іграшка_минула.xls" — містить "Артикул мережі" і "Нова АМ" '
                 'по магазинах за минулий період (стає полем «АМ»). Лист, що відповідає МТ і БШ, '
                 'визначається автоматично за збігом артикулів.',
            key="epicentr_reference",
        )
    with col6:
        price_file = st.file_uploader(
            "Прайс-лист з 1С УТП",
            type=FILE_TYPES,
            help='Наприклад "Прайс-лист <дата> Епік.xlsx" — той самий звіт «Прайс-лист (СКД)», що '
                 'і в Антошці, але сформований з мережевими кодами Епіцентру в колонці «Артикул в '
                 'сети». Дає «Артикул», «Бренд», «Статус артикула», «Склад» за збігом «Артикул '
                 'мережі» ↔ «Артикул в сети».',
            key="epicentr_price",
        )
    with col7:
        epicentr_delivery_file = st.file_uploader(
            "Графік поставок",
            type=FILE_TYPES,
            help='Наприклад "Графік поставок …xls". Дає «В дорозі» та кандидатів у новинки '
                 '(join за «Артикул», підтягнутим з прайс-листа).',
            key="epicentr_delivery",
        )

    epicentr_abc_files = st.file_uploader(
        "Файли АВС-аналізу продажів з 1С УТП (можна декілька)",
        type=FILE_TYPES,
        accept_multiple_files=True,
        help='Наприклад "ABC Chicco.xlsx" — звіт «АВС-аналіз продажів (за номенклатурою)», '
             'по одному файлу на кожну групу товарів. Додає поле «Категорія» («АВС-клас») — '
             'спільне для листів «МТ» і «БШ», співставляючи за «Артикул» (підтягнутим прайс-листом), '
             'а якщо не співпало — за точною «Назва».',
        key="epicentr_abc",
    )

    all_epicentr_files = [
        mt_season_file, mt_offseason_file,
        bsh_season_file, bsh_offseason_file,
        reference_file, price_file, epicentr_delivery_file, epicentr_abc_files,
    ]

    if st.button(
        "Сформувати звід",
        type="primary",
        disabled=not all(all_epicentr_files),
        key="epicentr_btn",
    ):
        try:
            with st.spinner("Обробляю файли..."):
                out_wb, all_stats = build_epicentr_combined_report(
                    mt_season_file, mt_offseason_file,
                    bsh_season_file, bsh_offseason_file,
                    reference_file,
                    price_file,
                    epicentr_delivery_file,
                    epicentr_abc_files,
                )

                buf = io.BytesIO()
                out_wb.save(buf)
                buf.seek(0)
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error(f"Не вдалося сформувати звід: {e}")
        else:
            st.success("Готово!")
            for label, stats in all_stats.items():
                st.markdown(f"**Лист «{label}»**")
                c1, c2, c3 = st.columns(3)
                c1.metric("Товарів", stats["n_products"])
                c2.metric("Магазинів", stats["n_stores"])
                c3.metric("Рядків товарів", stats["n_rows"])
                c4, c5, c6 = st.columns(3)
                c4.metric("Збігів з прайсом", stats["price_matches"])
                c5.metric("Збігів «В дорозі»", stats["delivery_matches"])
                c6.metric("Кандидатів у новинки", stats["n_novelty"])

                season_date = stats["season_date"] or "не розпізнано"
                offseason_date = stats["offseason_date"] or "не розпізнано"
                ref_info = (
                    f"довідник: лист «{stats['reference_sheet']}» ({stats['reference_overlap']} збігів для АМ)"
                    if stats["reference_sheet"]
                    else "довідник: не завантажений або відповідного листа не знайдено, «АМ» порожнє"
                )
                st.caption(
                    f"Дата «Кінцевий залишок на» — сезон: {season_date}, не сезон: {offseason_date}. "
                    f"«Залишок» взято з файлу **{stats['zalyshok_period']}**. {ref_info}."
                )

            out_name = f"Звід_Епіцентр_{Path(mt_season_file.name).stem}.xlsx"
            st.download_button(
                "Завантажити звід (.xlsx)",
                data=buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="epicentr_download",
            )

    st.divider()
    with st.expander("Що робить звід (коротко)"):
        st.markdown(
            """
- Один файл, два листи — «МТ» і «БШ», кожен по своїй парі файлів продажу.
- Один рядок на товар (артикул мережі) на листі, магазини — блоками колонок,
  підписаними кодом магазину (наприклад «01», «BR», «CHB»).
- **«Артикул мережі»** — 8-значний код товару з файлу продажу (єдиний
  ідентифікатор, який там є).
- **«Артикул»**, **«Бренд»**, **«Статус артикула»**, **«Склад»** —
  підтягуються з прайс-листа за збігом «Артикул мережі» ↔ «Артикул в сети»
  («Склад» = загальний залишок «Итого»). Якщо для конкретного товару збігу
  немає — поля порожні.
- **«В дорозі»** — рахується так само, як у Антошки, за щойно підтягнутим
  «Артикул».
- **«Категорія»** — підтягується з файлів АВС-аналізу («АВС-клас») за
  збігом «Артикул» (підтягнутим прайс-листом), а якщо не співпало — за
  точною «Назва» (без жодних відхилень). Спільне для листів «МТ» і «БШ» —
  можна завантажити декілька файлів АВС-аналізу відразу.
- **«АМ»** — підтягується з файлу-довідника за збігом «Артикул мережі»
  (лист довідника обирається автоматично за найбільшим перетином
  артикулів). Береться саме колонка **«Нова АМ»** довідника (а не стара
  «АМ») - те, що було «новим рішенням» минулого разу, зараз вже чинне.
  Якщо збігів немає, або «Нова АМ» там порожня - «АМ» лишається порожнім
  (без відкату до старої «АМ»).
- **«Сезон»** / **«Не сезон»** — кількість продажу («Розхід») з відповідного
  файлу.
- **«Залишок»** — береться з того файлу (сезон чи не сезон), у якого дата в
  заголовку «Кінцевий залишок на …» пізніша.
- **«НОВА АМ»** — поки порожнє поле зі світло-зеленою заливкою, заповнюється
  вручну.
- Кандидати в **новинки** — окремі рядки внизу кожного листа (лише «Назва»,
  «Артикул», «В дорозі»): та сама логіка, що в Антошки — шукаються тільки
  на листі Chicco графіка поставок, і лише якщо бренд Chicco є серед
  брендів, підтягнутих прайс-листом для цього листа («МТ» чи «БШ»).
- Рядки без назви товару (підсумки, заголовки магазинів) пропускаються.
"""
        )
