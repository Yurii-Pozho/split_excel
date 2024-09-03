import pandas as pd
import streamlit as st

# Інтерфейс Streamlit для завантаження файлу
st.title("Розділення аркушів Excel на окремі файли")

uploaded_file = st.file_uploader("Завантажте ваш Excel файл", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Завантажте всі аркуші з файлу
        sheets = pd.read_excel(uploaded_file, sheet_name=None)

        # Відобразити список аркушів
        st.write("Знайдено аркуші:", list(sheets.keys()))

        # Кнопка для збереження файлів
        if st.button("Зберегти кожен аркуш як окремий Excel файл"):
            for sheet_name, data in sheets.items():
                cleaned_name = sheet_name.replace('/', '')  # Просте очищення імені аркуша
                output_file = pd.ExcelWriter(f"{cleaned_name}.xlsx", engine='xlsxwriter')

                # Збереження файлу
                data.to_excel(output_file, index=False, sheet_name=cleaned_name)
                output_file.close()

                # Показати лінк для скачування файлу
                with open(f"{cleaned_name}.xlsx", "rb") as file:
                    st.download_button(
                        label=f"Скачати {cleaned_name}.xlsx",
                        data=file,
                        file_name=f"{cleaned_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
            st.success("Всі файли збережено!")

    except Exception as e:
        st.error(f"Сталася помилка: {e}")
