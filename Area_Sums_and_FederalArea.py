import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import ExcelWriter

st.title("📊 Генератор отчетов по регионам")
st.subheader("Загрузите Excel-файл для обработки")

# Блок загрузки файла
uploaded_file = st.file_uploader("Выберите файл .xlsx", type="xlsx")

if uploaded_file is not None:
    with st.spinner("Обработка файла..."):

        # Загрузка данных
        df = pd.read_excel(uploaded_file)

        # --- Первый отчет: по Области ---
        area_local = (
            df[df["Область"].notna() & (df["Область"] != "")]
            .groupby(["Область"])
            .agg({"Выкупили, шт.": "sum", "К перечислению за товар, руб.": "sum"})
            .astype(int)
            .reset_index()
        )
        area_local.sort_values(
            by="К перечислению за товар, руб.", ascending=False, inplace=True
        )

        # График для "Local_area"
        plt.figure(figsize=(12, 6))
        plt.bar(
            area_local["Область"].head(20),
            area_local["К перечислению за товар, руб."].head(20),
            color="skyblue",
        )
        plt.title("Сумма к перечислению по регионам (топ-20)")
        plt.xlabel("Регион")
        plt.ylabel("Сумма, руб.")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        chart_local_path = "chart_local.png"
        plt.savefig(chart_local_path)
        plt.close()

        # --- Второй отчет: по Федеральному округу ---
        area_federal = (
            df[df["Федеральный округ"].notna() & (df["Федеральный округ"] != "")]
            .groupby(["Федеральный округ"])
            .agg({"Выкупили, шт.": "sum", "К перечислению за товар, руб.": "sum"})
            .astype(int)
            .reset_index()
        )
        area_federal.sort_values(
            by="К перечислению за товар, руб.", ascending=False, inplace=True
        )

        # График для "Federal_area"
        plt.figure(figsize=(12, 6))
        plt.bar(
            area_federal["Федеральный округ"].head(10),
            area_federal["К перечислению за товар, руб."].head(10),
            color="lightgreen",
        )
        plt.title("Сумма к перечислению по федеральным округам")
        plt.xlabel("Федеральные округа")
        plt.ylabel("Сумма, руб.")
        plt.xticks(rotation=45, ha="right")
        plt.tight_layout()
        chart_federal_path = "chart_federal.png"
        plt.savefig(chart_federal_path)
        plt.close()

        # --- Сохранение в Excel ---
        output_path = "Sum_Area_Sveta.xlsx"

        with ExcelWriter(output_path, engine="openpyxl") as writer:
            # Лист 1: Local_area
            area_local.to_excel(writer, sheet_name="Local_area", index=False)
            ws1 = writer.sheets["Local_area"]
            img1 = XLImage(chart_local_path)
            ws1.add_image(img1, "F2")

            # Лист 2: Federal_area
            area_federal.to_excel(writer, sheet_name="Federal_area", index=False)
            ws2 = writer.sheets["Federal_area"]
            img2 = XLImage(chart_federal_path)
            ws2.add_image(img2, "F2")

        # --- Отправка файла пользователю ---
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Скачать готовый отчет",
                data=f,
                file_name=output_path,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.success("✅ Обработка завершена! Нажмите кнопку ниже для скачивания.")
else:
    st.info("📁 Пожалуйста, загрузите файл .xlsx для начала работы.")

# Очистка временных файлов после завершения
if os.path.exists("chart_local.png"):
    os.remove("chart_local.png")
if os.path.exists("chart_federal.png"):
    os.remove("chart_federal.png")
if os.path.exists("Sum_Area_Sveta.xlsx"):
    os.remove("Sum_Area_Sveta.xlsx")
