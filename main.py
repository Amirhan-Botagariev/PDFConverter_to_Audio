import fitz  # PyMuPDF
import edge_tts
import asyncio
import os
import shutil
from textwrap import wrap
from docx import Document

# Папка, где лежат файлы
BASE_FOLDER = "Async"

# Удаляемый текст
TEXT_TO_SKIP = "СКАЧАНО С WWW.SW.BAND - ПРИСОЕДИНЯЙСЯ!"

def clean_text(text):
    """Удаляет ненужные строки и короткие слова"""
    lines = text.split("\n")
    lines = [line.strip() for line in lines if TEXT_TO_SKIP not in line and line.strip()]
    cleaned_text = "\n".join(lines)
    return cleaned_text

async def text_to_speech(text, output_file, voice="ru-RU-SvetlanaNeural"):
    """Озвучивание текста только на русском с помощью Edge TTS"""
    text = clean_text(text)  # Убираем ненужный текст

    if not text.strip():
        print(f"Ошибка: Пустой текст в файле {output_file}!")
        return

    text_chunks = wrap(text, 4000)  # Разбиваем текст на части
    for idx, chunk in enumerate(text_chunks):
        chunk_file = output_file.replace(".mp3", f"_{idx}.mp3")
        tts = edge_tts.Communicate(chunk, voice)
        await tts.save(chunk_file)
        print(f"Сохранен аудиофайл: {chunk_file}")

def extract_text_from_pdf(pdf_path):
    """Извлекает текст из PDF и очищает его"""
    doc = fitz.open(pdf_path)
    text = "\n".join([page.get_text("text") for page in doc])
    return clean_text(text)

def extract_text_from_docx(docx_path):
    """Извлекает текст из DOCX и очищает его"""
    doc = Document(docx_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return clean_text(text)

def find_files(base_folder):
    """Ищет все PDF и DOCX файлы в папке и подпапках"""
    files = []
    for root, _, filenames in os.walk(base_folder):
        for filename in filenames:
            if filename.endswith(".pdf") or filename.endswith(".docx"):
                files.append(os.path.join(root, filename))
    return files

async def process_file(file_path):
    """Обрабатывает файл: извлекает текст, очищает его, сохраняет аудио, перемещает файл в папку"""
    file_name = os.path.basename(file_path)
    file_stem, ext = os.path.splitext(file_name)

    # Создаем папку для хранения результата
    file_output_folder = os.path.join(os.path.dirname(file_path), file_stem)
    os.makedirs(file_output_folder, exist_ok=True)

    # Извлекаем текст
    if ext.lower() == ".pdf":
        text = extract_text_from_pdf(file_path)
    elif ext.lower() == ".docx":
        text = extract_text_from_docx(file_path)
    else:
        return

    if not text.strip():  # Проверка на пустой текст после очистки
        print(f"Ошибка: Файл {file_name} не содержит полезного текста!")
        return

    # Создаем аудиофайл
    audio_file = os.path.join(file_output_folder, file_stem + ".mp3")
    await text_to_speech(text, audio_file)

    # Перемещаем обработанный файл в его новую папку
    new_file_path = os.path.join(file_output_folder, file_name)
    shutil.move(file_path, new_file_path)
    print(f"Файл {file_name} перемещен в {file_output_folder}")

async def main():
    files = find_files(BASE_FOLDER)
    if not files:
        print("Нет файлов для обработки!")
        return

    tasks = [process_file(file) for file in files]
    await asyncio.gather(*tasks)

asyncio.run(main())