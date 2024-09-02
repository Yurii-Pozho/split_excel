import pandas as pd
import streamlit as st
import re
import io
import warnings

# Ігнорувати попередження openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def clean_sheet_name(sheet_name):
    return re.sub(r'[\/:*?"<>|]', '', sheet_name)

# Інтерфейс Streamlit для завантаження файлу
st.title("Розділення аркушів Excel на окремі файли")

uploaded_file = st.file_uploader("Завантажте ваш Excel файл", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Завантажте всі аркуші з файлу
        sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')

        # Відобразити список аркушів
        st.write("Знайдено аркуші:", list(sheets.keys()))

        # Кнопка для збереження файлів
        if st.button("Зберегти кожен аркуш як окремий Excel файл"):
            for sheet_name, data in sheets.items():
                cleaned_name = clean_sheet_name(sheet_name)
                output_file = io.BytesIO()

                # Збереження файлу в пам'яті
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    data.to_excel(writer, index=False, sheet_name=cleaned_name)
                
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
