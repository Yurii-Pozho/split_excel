import pandas as pd
import streamlit as st
import os

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
            for i, (sheet_name, data) in enumerate(sheets.items()):
                # Зчитування значення з клітинки A1 для використання як назву файлу
                try:
                    file_name = data.iloc[0, 0]
                    if pd.isna(file_name) or not isinstance(file_name, str):
                        file_name = sheet_name  # Якщо A1 порожній або не рядок, використовується ім'я аркуша

                    file_name = file_name.replace('/', '')  # Очищення імені файлу
                    output_file = pd.ExcelWriter(f"{file_name}.xlsx", engine='xlsxwriter')

                    # Збереження файлу
                    data.to_excel(output_file, index=False, sheet_name=sheet_name)
                    output_file.close()

                    # Показати лінк для скачування файлу з унікальним ключем
                    with open(f"{file_name}.xlsx", "rb") as file:
                        st.download_button(
                            label=f"Скачати {file_name}.xlsx",
                            data=file,
                            file_name=f"{file_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{i}"  # Додання унікального ключа
                        )

                except Exception as e:
                    st.error(f"Сталася помилка з аркушем {sheet_name}: {e}")

            st.success("Всі файли збережено!")

    except Exception as e:
        st.error(f"Сталася помилка: {e}")
