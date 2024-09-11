import pandas as pd
import streamlit as st
import zipfile
import os
from io import BytesIO

# Інтерфейс Streamlit для завантаження файлу
st.title("Розділення аркушів Excel на окремі файли")

uploaded_file = st.file_uploader("Завантажте ваш Excel файл", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Завантажте всі аркуші з файлу без автоматичного визначення заголовків стовпців
        sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)

        # Кнопка для збереження файлів
        if st.button("Зберегти кожен аркуш як окремий Excel файл"):
            # Створення тимчасового архіву в пам'яті
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i, (sheet_name, data) in enumerate(sheets.items()):
                    try:
                        # Перевірка значення з клітинки A1 для імені файлу
                        base_name = str(data.iloc[0, 0]).strip() if not data.empty else sheet_name

                        # Перевірка значення з клітинки A6 для дати
                        if not data.empty and len(data) > 5:
                            date_value = data.iloc[5, 0]
                            # Перетворення на рядок дати у форматі YYYY-MM-DD
                            if pd.api.types.is_datetime64_any_dtype(pd.Series([date_value])):
                                date_str = pd.to_datetime(date_value).strftime('%Y-%m-%d')
                            else:
                                date_str = str(date_value).strip()  # Якщо не дата, використовуємо як є
                        else:
                            date_str = ""

                        # Об'єднання імені файлу з датою
                        file_name = f"{base_name}_{date_str}".replace('/', '').replace('\\', '')
                        file_name = file_name[:50]  # Обмеження довжини імені файлу до 50 символів
                        output_file_path = f"{file_name}.xlsx"

                        # Збереження аркушу у файлі Excel у пам'яті
                        excel_buffer = BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                            data.to_excel(writer, index=False, sheet_name=sheet_name)
                        excel_buffer.seek(0)

                        # Додавання файлу до ZIP-архіву
                        zip_file.writestr(output_file_path, excel_buffer.read())

                    except Exception as e:
                        st.error(f"Помилка з аркушем {sheet_name}: {e}")

            zip_buffer.seek(0)

            # Показати лінк для скачування ZIP-файлу
            st.download_button(
                label="Скачати всі файли як ZIP",
                data=zip_buffer,
                file_name="excel_sheets.zip",
                mime="application/zip"
            )

            st.success("Всі файли збережено в ZIP-архів!")
            

    except Exception as e:
        st.error(f"Помилка: {e}")
