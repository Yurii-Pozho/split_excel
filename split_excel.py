import pandas as pd
import streamlit as st
import re
from openpyxl import load_workbook
import io

# Інтерфейс Streamlit для завантаження файлу
st.title("Розділення аркушів Excel на окремі файли")

uploaded_file = st.file_uploader("Завантажте ваш Excel файл", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Завантажте всі аркуші з файлу
        sheets = pd.read_excel(uploaded_file, sheet_name=None)

        # Перейменування аркушів
        workbook = load_workbook(filename=io.BytesIO(uploaded_file.read()))
        for sheet in workbook.worksheets:
            # Отримайте значення з клітинки A1
            new_name = sheet['A1'].value

            if new_name:  # Перевірте, чи є значення
                # Очистіть назву аркуша (прибрати небажані символи)
                cleaned_name = re.sub(r'[\/:*?"<>|]', '', new_name)

                # Перейменуйте аркуш
                sheet.title = cleaned_name
        
        # Відобразити список аркушів
        st.write("Знайдено аркуші:", [sheet.title for sheet in workbook.worksheets])

        # Кнопка для збереження файлів
        if st.button("Зберегти кожен аркуш як окремий Excel файл"):
            for sheet in workbook.worksheets:
                cleaned_name = sheet.title
                output_file = io.BytesIO()

                # Збереження файлу в пам'яті
                with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                    df = pd.DataFrame(sheet.values)
                    df.to_excel(writer, index=False, sheet_name=cleaned_name)
                
                output_file.seek(0)

                # Показати лінк для скачування файлу
                st.download_button(
                    label=f"Скачати {cleaned_name}.xlsx",
                    data=output_file,
                    file_name=f"{cleaned_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            st.success("Всі файли збережено!")

    except Exception as e:
        st.error(f"Сталася помилка: {e}")
