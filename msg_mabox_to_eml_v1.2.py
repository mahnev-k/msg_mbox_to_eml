import sys
import os
import mimetypes
import extract_msg
import mailbox
import email
import subprocess
import logging
import re
import hashlib
import time
from pathlib import Path
from datetime import datetime
from email.utils import formatdate, parsedate_tz, mktime_tz
from email.header import decode_header

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,
    QFileDialog, QListWidget, QMessageBox, QProgressBar, QHBoxLayout,
    QToolButton, QMenu, QAction
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DragDropListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragEnabled(False)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QListWidget.DropOnly)
        self.setDefaultDropAction(Qt.CopyAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(('.msg', '.mbox')):
                    self.addItem(file_path)
            event.acceptProposedAction()


class ConversionWorker(QThread):
    """Отдельный поток для конвертации файлов"""
    progress = pyqtSignal(int)
    error = pyqtSignal(str, str)
    finished = pyqtSignal()  # Убираем параметр, всегда открываем каталог
    
    def __init__(self, files, output_dir, converter_instance):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.converter = converter_instance
    
    def run(self):
        total_files = len(self.files)
        
        for i, file_path in enumerate(self.files):
            try:
                if file_path.lower().endswith(".msg"):
                    self.converter.convert_msg_to_eml(file_path)
                elif file_path.lower().endswith(".mbox"):
                    self.converter.convert_mbox_to_eml(file_path)
                else:
                    continue
                    
            except Exception as e:
                logger.error(f"Ошибка при конвертации {file_path}: {str(e)}")
                self.error.emit(file_path, str(e))
                
            self.progress.emit(int((i + 1) / total_files * 100))
        
        self.finished.emit()


class MsgToEmlConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MSG и MBOX → EML Конвертер")
        self.setup_icon()
        self.resize(600, 450)
        self.output_dir = os.path.expanduser("~/EML_Export")
        self.conversion_worker = None
        
        self.init_ui()
        self.set_dark_theme()

    def setup_icon(self):
        """Безопасная установка иконки"""
        icon_paths = [
            os.path.expanduser("~/.config/msgtoeml/icon/msg2eml_gui.png"),
            "msg2eml_gui.png",  # В текущей директории
            "icon.png"
        ]
        
        for icon_path in icon_paths:
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                break
        else:
            # Используем системную иконку как fallback
            self.setWindowIcon(QIcon.fromTheme("mail-message-new"))

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)

        # Верхняя панель
        top_bar = QHBoxLayout()
        self.label = QLabel("Перетащите файлы .msg или .mbox, или выберите вручную:")
        top_bar.addWidget(self.label)

        self.settings_button = QToolButton()
        self.settings_button.setIcon(QIcon.fromTheme("preferences-system"))
        self.settings_button.setPopupMode(QToolButton.InstantPopup)
        self.setup_settings_menu()

        top_bar.addStretch()
        top_bar.addWidget(self.settings_button)
        layout.addLayout(top_bar)

        # Список файлов
        self.list_widget = DragDropListWidget()
        layout.addWidget(self.list_widget)

        # Кнопка очистки
        btn_layout = QHBoxLayout()
        self.clear_button = QPushButton("Очистить список")
        self.clear_button.clicked.connect(self.list_widget.clear)
        btn_layout.addStretch()
        btn_layout.addWidget(self.clear_button)
        layout.addLayout(btn_layout)

        # Кнопки выбора файлов
        self.select_button = QPushButton("Выбрать MSG-файлы")
        self.select_button.clicked.connect(self.select_files)
        layout.addWidget(self.select_button)

        self.mbox_button = QPushButton("Выбрать MBOX-файлы")
        self.mbox_button.clicked.connect(self.select_mbox_files)
        layout.addWidget(self.mbox_button)

        # Кнопка конвертации
        self.convert_button = QPushButton("Конвертировать")
        self.convert_button.clicked.connect(self.convert_all)
        layout.addWidget(self.convert_button)

        # Информация о папке вывода
        self.output_info = QLabel(f"Файлы будут сохранены в: {self.output_dir}")
        layout.addWidget(self.output_info)

        # Прогресс-бар
        self.progress = QProgressBar()
        self.progress.setTextVisible(True)
        self.progress.setAlignment(Qt.AlignCenter)
        self.progress.setFormat("%p%")
        layout.addWidget(self.progress)

        # Нижняя панель
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()
        self.mk_label = QLabel("MK")
        self.mk_label.setStyleSheet("color: gray; font-weight: bold;")
        self.mk_label.setTextInteractionFlags(Qt.NoTextInteraction)
        bottom_layout.addWidget(self.mk_label)
        layout.addLayout(bottom_layout)

    def setup_settings_menu(self):
        """Настройка меню настроек"""
        self.settings_menu = QMenu()

        self.theme_checkbox = QAction("Тёмная тема", self, checkable=True)
        self.theme_checkbox.setChecked(True)
        self.theme_checkbox.triggered.connect(self.toggle_theme)

        self.choose_dir_action = QAction("Выбрать папку для сохранения", self)
        self.choose_dir_action.triggered.connect(self.select_output_dir)

        self.settings_menu.addAction(self.theme_checkbox)
        self.settings_menu.addAction(self.choose_dir_action)
        self.settings_button.setMenu(self.settings_menu)

    def toggle_theme(self):
        """Переключение темы"""
        if self.theme_checkbox.isChecked():
            self.set_dark_theme()
        else:
            self.set_light_theme()

    def set_dark_theme(self):
        """Установка тёмной темы"""
        self.setStyleSheet("""
            QWidget { 
                background-color: #2b2b2b; 
                color: #ffffff; 
                font-family: Arial, sans-serif;
            }
            QPushButton { 
                background-color: #444; 
                padding: 8px; 
                border: 1px solid #666;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #555;
            }
            QPushButton:pressed {
                background-color: #333;
            }
            QListWidget { 
                background-color: #3b3b3b; 
                border: 1px solid #666;
                border-radius: 4px;
            }
            QProgressBar { 
                background-color: #555; 
                color: #fff; 
                border: 1px solid #666;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #0078d4;
                border-radius: 3px;
            }
        """)

    def set_light_theme(self):
        """Установка светлой темы"""
        self.setStyleSheet("")

    def decode_text(self, text):
        """Улучшенное декодирование текста с обработкой различных кодировок"""
        if isinstance(text, bytes):
            # Пробуем различные кодировки в порядке приоритета
            encodings = ['utf-8', 'windows-1251', 'cp1252', 'latin-1', 'ascii']
            for encoding in encodings:
                try:
                    return text.decode(encoding)
                except (UnicodeDecodeError, LookupError):
                    continue
            # Если ничего не сработало, используем replace
            return text.decode('utf-8', errors='replace')
        elif isinstance(text, str):
            return text
        else:
            return str(text or "")

    def decode_email_header(self, header_value):
        """Декодирование заголовков email с поддержкой MIME encoding"""
        if not header_value:
            return ""
        
        try:
            decoded_parts = decode_header(str(header_value))
            decoded_string = ""
            
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding:
                        try:
                            decoded_string += part.decode(encoding)
                        except (UnicodeDecodeError, LookupError):
                            decoded_string += part.decode('utf-8', errors='replace')
                    else:
                        decoded_string += part.decode('utf-8', errors='replace')
                else:
                    decoded_string += str(part)
                    
            return decoded_string
        except Exception as e:
            logger.warning(f"Ошибка декодирования заголовка '{header_value}': {str(e)}")
            return str(header_value)

    def parse_msg_date(self, date_value):
        """Улучшенный парсинг даты из MSG файла"""
        if not date_value:
            return datetime.now()
        
        try:
            # Если это строка, декодируем её
            if isinstance(date_value, str):
                decoded_date = self.decode_email_header(date_value)
                # Удаляем комментарии в скобках вроде (*31.12.1899 05:31:40*)
                decoded_date = re.sub(r'\s*\([^)]*\)\s*', '', decoded_date).strip()
                
                # Попробуем распарсить как email дату
                try:
                    parsed_tuple = parsedate_tz(decoded_date)
                    if parsed_tuple:
                        timestamp = mktime_tz(parsed_tuple)
                        return datetime.fromtimestamp(timestamp)
                except (ValueError, TypeError, OverflowError):
                    pass
            
            # Если это объект datetime
            elif hasattr(date_value, 'year'):
                # Проверяем на некорректную дату (1899 год часто означает ошибку)
                if date_value.year < 1970:
                    logger.warning(f"Некорректная дата: {date_value}, используем текущую")
                    return datetime.now()
                return date_value
            
            # Если это timestamp
            elif isinstance(date_value, (int, float)):
                if date_value > 0:
                    return datetime.fromtimestamp(date_value)
            
        except Exception as e:
            logger.warning(f"Ошибка парсинга даты '{date_value}': {str(e)}")
        
        # Fallback - возвращаем текущую дату
        return datetime.now()

    def get_safe_recipients(self, recipients):
        """Безопасное получение списка получателей"""
        if not recipients:
            return ""
        
        recipient_list = []
        for recipient in recipients:
            try:
                email_addr = getattr(recipient, 'email', None)
                display_name = getattr(recipient, 'display_name', None)
                
                if email_addr:
                    recipient_list.append(email_addr)
                elif display_name:
                    recipient_list.append(display_name)
                else:
                    recipient_list.append("Unknown Recipient")
            except AttributeError:
                recipient_list.append("Unknown Recipient")
        
        return ", ".join(recipient_list)

    def select_files(self):
        """Выбор MSG файлов"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите файлы", "", "MSG файлы (*.msg)"
        )
        for f in files:
            self.list_widget.addItem(f)

    def select_mbox_files(self):
        """Выбор MBOX файлов"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите MBOX-файлы", "", "MBOX файлы (*.mbox)"
        )
        for f in files:
            self.list_widget.addItem(f)

    def select_output_dir(self):
        """Выбор папки для сохранения"""
        dir_path = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if dir_path:
            self.output_dir = dir_path
            self.output_info.setText(f"Файлы будут сохранены в: {self.output_dir}")

    def convert_all(self):
        """Запуск конвертации всех файлов"""
        count = self.list_widget.count()
        if count == 0:
            QMessageBox.warning(self, "Ошибка", "Не выбраны файлы")
            return

        # Создаем папку вывода если не существует
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)

        # Собираем список файлов
        files = [self.list_widget.item(i).text() for i in range(count)]

        # Отключаем кнопку конвертации
        self.convert_button.setEnabled(False)
        self.convert_button.setText("Конвертация...")

        # Запускаем конвертацию в отдельном потоке
        self.conversion_worker = ConversionWorker(files, self.output_dir, self)
        self.conversion_worker.progress.connect(self.progress.setValue)
        self.conversion_worker.error.connect(self.handle_conversion_error)
        self.conversion_worker.finished.connect(self.handle_conversion_finished)
        self.conversion_worker.start()

    def handle_conversion_error(self, file_path, error_message):
        """Обработка ошибок конвертации"""
        QMessageBox.critical(self, "Ошибка", f"Файл: {file_path}\nОшибка: {error_message}")

    def handle_conversion_finished(self):
        """Завершение конвертации"""
        self.list_widget.clear()
        self.progress.setValue(100)
        
        # Включаем кнопку обратно
        self.convert_button.setEnabled(True)
        self.convert_button.setText("Конвертировать")
        
        QMessageBox.information(self, "Готово", "Конвертация завершена!")
        
        # Всегда открываем каталог с результатами
        self.open_result(self.output_dir)

    def open_result(self, path):
        """Безопасное открытие файла или папки"""
        try:
            if sys.platform.startswith("linux"):
                subprocess.Popen(["xdg-open", path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif sys.platform == "win32":
                os.startfile(path)
        except Exception as e:
            logger.error(f"Не удалось открыть {path}: {str(e)}")
            QMessageBox.warning(
                self, "Предупреждение", 
                f"Не удалось открыть файл или папку:\n{str(e)}"
            )

    def is_inline_attachment(self, att):
        """Определяет, является ли вложение встроенным изображением"""
        filename = (getattr(att, 'longFilename', None) or 
                   getattr(att, 'shortFilename', None) or "")
        
        # Проверяем расширение файла
        if not filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
            return False
        
        # Проверяем наличие Content-ID (cid)
        cid = getattr(att, 'cid', None) or getattr(att, 'contentId', None)
        if cid:
            return True
            
        # Дополнительные проверки для встроенных изображений
        # Иногда встроенные изображения имеют специальные атрибуты
        attachment_method = getattr(att, 'attachmentMethod', None)
        if attachment_method == 6:  # ATTACH_EMBEDDED_MSG в некоторых случаях
            return False
        
        # Проверяем по имени файла - часто встроенные изображения имеют автоматические имена
        if re.match(r'^image\d+\.(png|jpg|jpeg|gif|bmp)$', filename.lower()):
            return True
            
        return False

    def process_html_with_inline_images(self, html_body, attachments, cid_mapping):
        """Обрабатывает HTML тело письма, заменяя ссылки на изображения"""
        if not html_body or not attachments:
            return html_body
        
        # Ищем все ссылки на изображения в HTML
        img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
        
        def replace_img_src(match):
            full_tag = match.group(0)
            src = match.group(1)
            
            # Если это уже cid: ссылка, оставляем как есть
            if src.startswith('cid:'):
                return full_tag
            
            # Ищем соответствующее изображение по имени файла
            for att in attachments:
                filename = (getattr(att, 'longFilename', None) or 
                           getattr(att, 'shortFilename', None) or "")
                
                if filename and (filename in src or src.endswith(filename)):
                    # Генерируем Content-ID для этого изображения
                    cid = f"img_{hashlib.md5(filename.encode()).hexdigest()[:8]}"
                    cid_mapping[filename] = cid
                    
                    # Заменяем src на cid
                    new_tag = re.sub(r'src=["\'][^"\']+["\']', f'src="cid:{cid}"', full_tag)
                    return new_tag
            
            return full_tag
        
        return re.sub(img_pattern, replace_img_src, html_body, flags=re.IGNORECASE)

    def convert_msg_to_eml(self, msg_path):
        """Конвертация MSG файла в EML с правильной обработкой встроенных изображений"""
        try:
            msg = extract_msg.Message(msg_path)
            
            # Безопасное получение данных сообщения
            msg_sender = getattr(msg, 'sender', None) or ""
            msg_to = self.get_safe_recipients(getattr(msg, 'recipients', None))
            msg_subject = getattr(msg, 'subject', None) or ""
            msg_body = self.decode_text(getattr(msg, 'body', None))
            msg_html = self.decode_text(getattr(msg, 'htmlBody', None))

            # Обработка вложений
            attachments = getattr(msg, 'attachments', [])
            inline_attachments = []
            regular_attachments = []
            cid_mapping = {}  # Маппинг имени файла -> Content-ID

            # Разделяем вложения на встроенные и обычные
            for att in attachments:
                if self.is_inline_attachment(att):
                    inline_attachments.append(att)
                else:
                    regular_attachments.append(att)

            # Обрабатываем HTML с встроенными изображениями
            if msg_html and inline_attachments:
                msg_html = self.process_html_with_inline_images(msg_html, inline_attachments, cid_mapping)

            # Создание структуры письма
            if not msg_body and not msg_html:
                # Простая структура для пустых сообщений
                outer = email.mime.text.MIMEText("", "plain", "utf-8")
            elif msg_html and not msg_body:
                # Только HTML контент
                if inline_attachments:
                    # HTML с встроенными изображениями
                    outer = MIMEMultipart("related")
                    html_part = MIMEText(msg_html, "html", "utf-8")
                    outer.attach(html_part)
                    
                    # Добавляем встроенные изображения
                    for att in inline_attachments:
                        self.process_inline_attachment(att, outer, cid_mapping)
                else:
                    outer = MIMEText(msg_html, "html", "utf-8")
            elif msg_body and not msg_html:
                # Только текстовый контент
                outer = MIMEText(msg_body, "plain", "utf-8")
            else:
                # И текст и HTML
                if inline_attachments:
                    # Создаем сложную структуру: alternative с related для HTML
                    outer = MIMEMultipart("alternative")
                    
                    # Добавляем текстовую часть
                    text_part = MIMEText(msg_body, "plain", "utf-8")
                    outer.attach(text_part)
                    
                    # Создаем related контейнер для HTML с изображениями
                    html_related = MIMEMultipart("related")
                    html_part = MIMEText(msg_html, "html", "utf-8")
                    html_related.attach(html_part)
                    
                    # Добавляем встроенные изображения
                    for att in inline_attachments:
                        self.process_inline_attachment(att, html_related, cid_mapping)
                    
                    outer.attach(html_related)
                else:
                    # Простая alternative структура
                    outer = MIMEMultipart("alternative")
                    outer.attach(MIMEText(msg_body, "plain", "utf-8"))
                    outer.attach(MIMEText(msg_html, "html", "utf-8"))

            # Если есть обычные вложения, оборачиваем все в mixed
            if regular_attachments:
                if isinstance(outer, MIMEMultipart):
                    mixed_outer = MIMEMultipart("mixed")
                    # Копируем заголовки
                    for key, value in outer.items():
                        mixed_outer[key] = value
                    mixed_outer.attach(outer)
                else:
                    mixed_outer = MIMEMultipart("mixed")
                    # Копируем заголовки
                    for key, value in outer.items():
                        mixed_outer[key] = value
                    mixed_outer.attach(outer)
                
                outer = mixed_outer

                # Добавляем обычные вложения
                for att in regular_attachments:
                    try:
                        self.process_regular_attachment(att, outer)
                    except Exception as e:
                        logger.warning(f"Ошибка при обработке вложения: {str(e)}")
                        continue

            # Обязательные заголовки для Evolution
            outer["Subject"] = msg_subject
            outer["From"] = msg_sender
            outer["To"] = msg_to
            outer["Message-ID"] = f"<{hash(msg_path)}@converted.local>"
            
            # Исправленная обработка даты
            msg_date = None
            if hasattr(msg, 'date') and msg.date:
                msg_date = self.parse_msg_date(msg.date)
            else:
                # Пытаемся найти дату в других полях
                for attr in ['creationTime', 'lastModificationTime', 'receivedTime']:
                    if hasattr(msg, attr):
                        attr_value = getattr(msg, attr)
                        if attr_value:
                            msg_date = self.parse_msg_date(attr_value)
                            break

            if not msg_date:
                msg_date = datetime.now()

            # Форматируем дату для email заголовка
            try:
                formatted_date = formatdate(msg_date.timestamp(), localtime=True)
            except:
                formatted_date = formatdate(time.time(), localtime=True)

            outer["Date"] = formatted_date

            # MIME версия обязательна
            outer["MIME-Version"] = "1.0"

            # Сохранение файла
            out_filename = self.generate_safe_filename(msg_path, ".eml")
            out_path = os.path.join(self.output_dir, out_filename)
            
            # Сохраняем с правильными заголовками для Unix
            with open(out_path, "w", encoding="utf-8", newline='\n') as f:
                f.write(outer.as_string())

            logger.info(f"Успешно конвертирован: {msg_path} -> {out_path}")
            return out_path

        except Exception as e:
            logger.error(f"Ошибка конвертации MSG файла {msg_path}: {str(e)}")
            raise

    def process_inline_attachment(self, att, parent, cid_mapping):
        """Обработка встроенных изображений с правильным Content-ID"""
        filename = (getattr(att, 'longFilename', None) or 
                   getattr(att, 'shortFilename', None) or 
                   "inline_image")
        
        data = getattr(att, 'data', None)
        if data is None:
            return
            
        # Нормализация данных
        if isinstance(data, str):
            data = data.encode(errors="replace")
        elif not isinstance(data, bytes):
            data = bytes(data)

        # Определение MIME типа
        mime_type, _ = mimetypes.guess_type(filename)
        if mime_type and mime_type.startswith('image/'):
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "image", "png"

        # Создаем изображение
        if maintype == "image":
            attachment = MIMEImage(data, subtype)
        else:
            attachment = MIMEBase(maintype, subtype)
            attachment.set_payload(data)
            encoders.encode_base64(attachment)

        # Устанавливаем Content-ID
        if filename in cid_mapping:
            cid = cid_mapping[filename]
        else:
            # Генерируем новый Content-ID
            existing_cid = getattr(att, 'cid', None) or getattr(att, 'contentId', None)
            if existing_cid:
                cid = existing_cid.strip('<>')
            else:
                cid = f"img_{hashlib.md5(filename.encode()).hexdigest()[:8]}"
        
        attachment.add_header("Content-ID", f"<{cid}>")
        attachment.add_header("Content-Disposition", "inline", filename=filename)
        
        parent.attach(attachment)

    def process_regular_attachment(self, att, outer):
        """Обработка обычных вложений"""
        filename = (getattr(att, 'longFilename', None) or 
                   getattr(att, 'shortFilename', None) or 
                   "attachment")
        
        data = getattr(att, 'data', None)
        if data is None:
            return
            
        # Нормализация данных
        if isinstance(data, str):
            data = data.encode(errors="replace")
        elif not isinstance(data, bytes):
            data = bytes(data)

        # Определение MIME типа
        mime_type, _ = mimetypes.guess_type(filename)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        # Создаем обычное вложение
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(data)
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", "attachment", filename=filename)
        attachment.add_header("Content-Transfer-Encoding", "base64")
        outer.attach(attachment)

    def convert_mbox_to_eml(self, mbox_path):
        """Конвертация MBOX файла в EML с улучшенной обработкой ошибок"""
        out_path = None
        converted_count = 0
        
        try:
            mbox = mailbox.mbox(mbox_path)
            base_name = os.path.splitext(os.path.basename(mbox_path))[0]
            parent_dir = os.path.basename(os.path.dirname(mbox_path))
            prefix = f"{parent_dir}_{base_name}" if parent_dir else base_name
            
            for i, message in enumerate(mbox):
                try:
                    # Получаем сырые байты сообщения
                    raw_bytes = message.as_bytes()
                    
                    # Парсим сообщение
                    parsed = email.message_from_bytes(raw_bytes)
                    
                    # Убеждаемся что есть обязательные заголовки
                    if not parsed.get("Message-ID"):
                        parsed["Message-ID"] = f"<{hash(raw_bytes)}@converted.local>"
                    
                    if not parsed.get("Date"):
                        parsed["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z")
                    
                    if not parsed.get("MIME-Version"):
                        parsed["MIME-Version"] = "1.0"
                    
                    out_filename = f"{prefix}_{i + 1:04d}.eml"
                    out_path = os.path.join(self.output_dir, out_filename)
                    
                    # Сохраняем с правильным форматом строк для Unix
                    with open(out_path, "w", encoding="utf-8", newline='\n') as f:
                        f.write(parsed.as_string())
                    
                    converted_count += 1
                    
                except Exception as e:
                    logger.warning(f"Ошибка при обработке сообщения {i + 1} из {mbox_path}: {str(e)}")
                    continue
            
            if converted_count > 0:
                logger.info(f"Конвертировано {converted_count} сообщений из {mbox_path}")
            else:
                logger.warning(f"Не удалось конвертировать ни одного сообщения из {mbox_path}")
                
        except Exception as e:
            logger.error(f"Ошибка при открытии MBOX файла {mbox_path}: {str(e)}")
            raise
            
        return out_path

    def generate_safe_filename(self, original_path, extension):
        """Генерация безопасного имени файла"""
        base_name = os.path.splitext(os.path.basename(original_path))[0]
        # Удаляем небезопасные символы
        safe_name = "".join(c for c in base_name if c.isalnum() or c in (' ', '-', '_')).strip()
        if not safe_name:
            safe_name = "converted_file"
        return safe_name + extension

    def closeEvent(self, event):
        """Обработка закрытия приложения"""
        if self.conversion_worker and self.conversion_worker.isRunning():
            reply = QMessageBox.question(
                self, 'Закрытие приложения',
                'Конвертация в процессе. Вы действительно хотите закрыть приложение?',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.conversion_worker.terminate()
                self.conversion_worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName("MSG to EML Converter")
    app.setApplicationVersion("2.0")
    
    window = MsgToEmlConverter()
    window.show()
    
    sys.exit(app.exec_())
