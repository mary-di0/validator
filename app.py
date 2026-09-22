import io
import shutil
import zipfile
import tempfile
from pathlib import Path

import streamlit as st
import valid


st.set_page_config(page_title="Валидация документов", layout="wide")
st.title("Валидация документов базы знаний")
st.write("Загрузите документы (или ZIP-архив с папкой) и нажмите «Запустить валидацию».")


# ---------- 1. Загрузка файлов ----------
uploaded_files = st.file_uploader(
    "Выберите файлы (можно несколько) или ZIP-архив",
    type=["docx", "doc", "pdf", "txt", "md", "zip", "xlsx"],
    accept_multiple_files=True,
)

run_button = st.button("Запустить валидацию", type="primary")


if run_button and uploaded_files:
    # ---------- 2. Готовим временную папку ----------
    work_dir = Path(tempfile.mkdtemp(prefix="valid_"))
    docs_dir = work_dir / "documents"
    docs_dir.mkdir(parents=True, exist_ok=True)

    with st.status("Распаковка файлов…", expanded=True) as status:
        for uf in uploaded_files:
            if uf.name.lower().endswith(".zip"):
                # распаковываем архив в docs_dir
                with zipfile.ZipFile(io.BytesIO(uf.read())) as z:
                    for member in z.namelist():
                        # защита от zip-slip
                        target = (docs_dir / member).resolve()
                        if not str(target).startswith(str(docs_dir.resolve())):
                            continue
                        z.extract(member, docs_dir)
            else:
                (docs_dir / uf.name).write_bytes(uf.read())
        status.update(label=f"Файлы готовы: {len(list(docs_dir.rglob('*')))} объектов", state="complete")

    # ---------- 3. Прогон пайплайна ----------
    output_xlsx = work_dir / "Отчет_валидации.xlsx"

    with st.status("Анализ документов…", expanded=True) as status:
        # Подменяем конфиг и схему как в main()
        valid.Config.INPUT_DOCS_DIR = docs_dir
        valid.Config.OUTPUT_REPORT = output_xlsx
        valid.Config.TEST_MODE = False

        # API-параметры — можно вынести в поля UI или в secrets
        valid.Config.API_URL = st.secrets.get("API_URL", valid.Config.API_URL)
        valid.Config.MODEL = st.secrets.get("MODEL", valid.Config.MODEL)
        valid.Config.TOKEN = st.secrets.get("TOKEN", valid.Config.TOKEN)

        valid.SCHEMA = valid.load_schema("schema.json")
        if valid.SCHEMA is None:
            st.error("Не загружена схема schema.json")
            st.stop()

        # ШАГ 1 — анализ документов
        st.write("Шаг 1: анализ документов…")
        (report_rows, file_texts, file_paths,
         verdicts_by_file, error_rows) = valid.process_all_documents(docs_dir)

        # ШАГ 2 — дубли и версии
        st.write("Шаг 2: поиск дублей и версий…")
        pairs = valid.compute_similarity_pairs(file_texts)
        duplicate_rows = valid.find_cross_file_duplicates(file_texts, pairs=pairs)
        version_groups = valid.find_version_groups(file_paths, pairs, threshold_low=0.5)
        version_rows = valid.build_version_group_rows(version_groups, file_paths, verdicts_by_file)


        files_in_version_groups = {name: g for g in version_groups for name in g}
        for row in report_rows:
            if row["Критерий"].startswith("2.4") and row["Файл"] in files_in_version_groups:
                group = files_in_version_groups[row["Файл"]]
                others = [f for f in group if f != row["Файл"]]
                note = (
                    f"В папке documents найдены другие файлы, похожие на этот документ "
                    f"(вероятно, другие версии/копии): {', '.join(others)}. По правилам "
                    f"в базе знаний должна храниться только 1 актуальная версия, "
                    f"остальные должны быть архивированы в другом месте (см. лист "
                    f"'Версии документов')."
                )
                row["Статус"] = "Есть нарушения"
                row["Находки"] = (
                    (row["Находки"] + "; " if row["Находки"] and row["Находки"] != "—" else "")
                    + note
                )

        # ШАГ 3 — сохранение
        st.write("Шаг 3: сохранение отчёта…")
        valid.save_report(report_rows, duplicate_rows, version_rows, error_rows, output_xlsx)

        status.update(label="Готово!", state="complete")

    # ---------- 4. Краткая сводка ----------
    col1, col2, col3 = st.columns(3)
    col1.metric("Файлов обработано", len(file_texts))
    col2.metric("Дублей/похожих пар", len(duplicate_rows))
    col3.metric("Групп версий", len(version_groups))

    if error_rows:
        with st.expander(f"Ошибок обработки: {len(error_rows)}"):
            st.dataframe(error_rows)

    # ---------- 5. Кнопка скачивания ----------
    with open(output_xlsx, "rb") as f:
        st.download_button(
            label="Скачать отчёт Excel",
            data=f.read(),
            file_name="Отчет_валидации.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )