import os
import pandas as pd
import folium
from geopy.geocoders import Nominatim
from time import sleep, time
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading

# Глобальные переменные
output_map_file = None
uploaded_file_path = None
missing_coords_file = "missing_coordinates.txt"  # Имя файла для записи недостающих координат

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Функция для записи недостающих координат в файл
def write_missing_coordinates(city, region):
    with open(missing_coords_file, 'a') as file:
        file.write(f"{city}, {region}\n")


# Функция для запуска процесса обработки карты
def start_processing():
    global output_map_file, uploaded_file_path, missing_coords_file
    try:
        if not uploaded_file_path or not os.path.exists(uploaded_file_path):
            raise FileNotFoundError("Не найден файл для обработки. Пожалуйста, загрузите файл.")

        # Очистка файла с недостающими координатами перед началом новой обработки
        if os.path.exists(missing_coords_file):
            os.remove(missing_coords_file)

        # Шаг 1: Чтение файла Excel
        logging.info("Чтение файла Excel...")
        df = pd.read_excel(uploaded_file_path)
        logging.info(f"Файл успешно загружен. Общее количество строк: {df.shape[0]}")

        # Шаг 2: Фильтрация данных и удаление дубликатов
        logging.info("Фильтрация данных и удаление дубликатов...")
        df_filtered = df[['Area', 'City', 'Доставка']].drop_duplicates()
        logging.info(f"Количество уникальных строк: {df_filtered.shape[0]}")

        # Шаг 3: Инициализация карты Украины
        logging.info("Инициализация карты Украины...")
        map_ukraine = folium.Map(location=[48.3794, 31.1656], zoom_start=6)

        # Шаг 4: Инициализация геокодера
        logging.info("Инициализация геокодера...")
        geolocator = Nominatim(user_agent="ukraine_map")

        # Прогресс-бар
        progress_bar['value'] = 0
        progress_bar['maximum'] = len(df_filtered)

        start_time = time()  # Время начала обработки

        # Шаг 5: Добавление кружочков на карту
        logging.info("Добавление кружочков на карту...")
        for i, row in enumerate(df_filtered.iterrows()):
            lat, lon = get_coordinates(geolocator, row[1]['City'], row[1]['Area'])
            if lat and lon:
                color = 'green' if row[1]['Доставка'].strip().lower() == 'да' else 'red'
                folium.CircleMarker(
                    location=[lat, lon],
                    radius=5,  # Радиус кружочка
                    color=color,
                    fill=True,
                    fill_color=color,
                    fill_opacity=0.7,
                    popup=f"{row[1]['City']} ({row[1]['Area']}): {row[1]['Доставка']}"
                ).add_to(map_ukraine)
            else:
                logging.warning(f"Не удалось найти координаты для {row[1]['City']}, {row[1]['Area']}")
                write_missing_coordinates(row[1]['City'], row[1]['Area'])  # Записываем недостающие координаты

            # Обновляем прогресс-бар и время
            progress_bar['value'] = i + 1
            progress_bar.update()
            elapsed_time = time() - start_time
            estimated_total_time = (elapsed_time / (i + 1)) * len(df_filtered)
            remaining_time = estimated_total_time - elapsed_time
            time_label.config(text=f"Время работы: {int(elapsed_time)} сек. | Осталось: {int(remaining_time)} сек.")
            sleep(0.5)  # Задержка для избегания блокировки геокодера

        # Шаг 6: Сохранение карты в файл
        output_map_file = 'ukraine_map_circles_full.html'
        logging.info(f"Сохранение карты в файл {output_map_file}...")
        map_ukraine.save(output_map_file)
        logging.info("Карта успешно сохранена.")
        messagebox.showinfo("Завершено",
                            "Карта успешно сгенерирована. Чтобы сохранить карту и файл с недостающими координатами, нажмите 'Сохранить файл'.")
        save_button.config(state="normal")
        time_label.config(text="Обработка завершена.")
    except Exception as e:
        logging.error(f"Ошибка: {e}")
        messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")


# Функция для очистки папки uploads
def clean_uploads_directory():
    uploads_dir = 'uploads'
    for file in os.listdir(uploads_dir):
        file_path = os.path.join(uploads_dir, file)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                logging.info(f"Удален файл: {file_path}")
        except Exception as e:
            logging.error(f"Ошибка при удалении файла {file_path}: {e}")


# Функция для загрузки файла
def upload_file():
    global uploaded_file_path
    file_path = filedialog.askopenfilename(title="Выберите Excel файл", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        file_name = os.path.basename(file_path)
        target_path = os.path.join('uploads', file_name)
        os.makedirs('uploads', exist_ok=True)
        os.replace(file_path, target_path)
        uploaded_file_path = target_path  # Сохраняем путь загруженного файла
        messagebox.showinfo("Загрузка завершена", f"Файл загружен в папку uploads как {file_name}")
        process_button.config(state="normal")


# Функция для получения координат
def get_coordinates(geolocator, city, region):
    try:
        location = geolocator.geocode(f"{city}, {region}, Ukraine")
        if location:
            return location.latitude, location.longitude
        else:
            return None, None
    except Exception as e:
        logging.error(f"Ошибка при геокодировании {city}, {region}: {e}")
        return None, None


# Функция для сохранения итогового HTML файла и файла с недостающими координатами
def save_file():
    global output_map_file, missing_coords_file
    if output_map_file and os.path.exists(output_map_file):
        save_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")])
        if save_path:
            os.replace(output_map_file, save_path)
            messagebox.showinfo("Файл сохранен", f"Карта успешно сохранена как {save_path}")

        # Сохранение файла с недостающими координатами
        if os.path.exists(missing_coords_file):
            save_txt_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
            if save_txt_path:
                os.replace(missing_coords_file, save_txt_path)
                messagebox.showinfo("Файл сохранен",
                                    f"Файл с недостающими координатами успешно сохранен как {save_txt_path}")

        output_map_file = None
        save_button.config(state="disabled")
    else:
        messagebox.showerror("Ошибка", "Файл для сохранения не найден.")


# Создание основного окна
root = tk.Tk()
root.title("Генератор карты Украины")
root.geometry("400x350")  # Увеличим размер окна для добавления времени

# Кнопка для загрузки файла
upload_button = tk.Button(root, text="Загрузить файл", command=upload_file)
upload_button.pack(pady=10)

# Прогресс-бар
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Метка для отображения времени
time_label = tk.Label(root, text="Ожидание начала обработки...")
time_label.pack(pady=10)

# Кнопка для обработки файла
process_button = tk.Button(root, text="Обработать файл",
                           command=lambda: threading.Thread(target=start_processing).start(), state="disabled")
process_button.pack(pady=10)

# Кнопка для сохранения файла
save_button = tk.Button(root, text="Сохранить файл", command=save_file, state="disabled")
save_button.pack(pady=10)

# Запуск главного цикла приложения
root.mainloop()
