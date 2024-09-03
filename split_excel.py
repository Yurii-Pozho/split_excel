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
                try:
                    # Зчитування значення з клітинки A1 для використання як назву файлу
                    file_name = str(data.iloc[0, 0]).strip()
                    if pd.isna(file_name) or not file_name:
                        file_name = sheet_name  # Якщо A1 порожній або не рядок, використовується ім'я аркуша
                    
                    # Очищення імені файлу
                    file_name = file_name.replace('/', '').replace('\\', '')  # Видалення символів / і \
                    file_name = file_name[:50]  # Обмеження довжини імені файлу до 50 символів
                    
                    # Збереження файлу
                    output_file_path = f"{file_name}.xlsx"
                    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as output_file:
                        data.to_excel(output_file, index=False, sheet_name=sheet_name)

                    # Показати лінк для скачування файлу з унікальним ключем
                    with open(output_file_path, "rb") as file:
                        st.download_button(
                            label=f"Скачати {file_name}.xlsx",
                            data=file,
                            file_name=output_file_path,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"download_{i}"  # Додання унікального ключа
                        )

                except Exception as e:
                    st.error(f"Сталася помилка з аркушем {sheet_name}: {e}")

            st.success("Всі файли збережено!")

    except Exception as e:
        st.error(f"Сталася помилка: {e}")
