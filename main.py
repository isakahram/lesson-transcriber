#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
«Транскрипция урока» — простая программа для записи онлайн-уроков
и получения текстовой расшифровки (транскрипции) в txt-файл.

Как это работает:
  1. Программа одновременно записывает два звуковых потока:
     - системный звук (то, что слышно в колонках/наушниках: голос ученика в Zoom, браузер);
     - микрофон (голос репетитора).
  2. После нажатия «Остановить» две дорожки смешиваются в одну.
  3. Речь распознаётся локально моделью Whisper (библиотека faster-whisper).
     Интернет нужен только один раз — чтобы скачать модель. Никаких API-ключей
     и облачных сервисов не используется.
  4. Готовый текст сохраняется на рабочий стол в файл «Занятие ГГГГ-ММ-ДД.txt».

Вся программа находится в одном файле — так её проще запускать и править.
Порядок разделов:
  1. Настройки
  2. Вспомогательные функции (ошибки, пути, работа с аудио)
  3. Захват звука (системный звук + микрофон)
  4. Распознавание речи (faster-whisper)
  5. Окно программы (tkinter)
  6. Точка входа
"""

from __future__ import annotations

# ============================================================
#  ИМПОРТЫ
# ============================================================

import ctypes             # нужно, чтобы узнать путь к «Рабочему столу» и инициализировать COM
import os
import queue
import re
import shutil
import sys
import tempfile
import threading
import time
import traceback
from datetime import datetime

# numpy и soundfile нужны всегда, поэтому проверяем их сразу и, если чего-то нет,
# печатаем понятную инструкцию вместо длинной простыни ошибок.
try:
    import numpy as np
    import soundfile as sf
except ImportError as exc:  # pragma: no cover — зависит от настроек компьютера
    print(
        "Не установлена нужная библиотека: "
        f"{exc.name}.\n\n"
        "Откройте командную строку в папке с программой и выполните команду:\n"
        "    pip install -r requirements.txt\n\n"
        "Подробная инструкция — в файле README.md."
    )
    raise SystemExit(1)

# tkinter встроен в Python (в установщике с python.org он есть по умолчанию),
# но на всякий случай не даём программе «упасть» непонятной ошибкой.
try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except ImportError:  # pragma: no cover — редкий случай
    tk = None
    ttk = None
    messagebox = None


# ============================================================
#  1. НАСТРОЙКИ (можно менять, но обычно этого не требуется)
# ============================================================

SAMPLE_RATE = 16000          # частота дискретизации для Whisper — только 16000 Гц
MODEL_SIZE = "small"         # размер модели Whisper: "small" — хороший баланс.
#                              Если компьютер очень медленный, можно поставить "base".
COMPUTE_TYPE = "int8"        # ускоренное вычисление на процессоре
LANGUAGE = "ru"              # язык урока
BEAM_SIZE = 3                # качество распознавания (больше — точнее и медленнее)
VAD_MIN_SILENCE_MS = 500     # тишина короче этого времени не разрывает фразы

RECORD_BLOCK_SECONDS = 0.1   # читаем звук блоками по 0,1 секунды
TRANSCRIBE_CHUNK_SECONDS = 300   # длинные уроки распознаём кусками по 5 минут
BOUNDARY_SEARCH_SECONDS = 15     # границу куска ищем в пределах ±15 секунд
MIX_BLOCK_FRAMES = 1 << 18       # размер блока при смешивании (262144 отсчёта)
RESAMPLE_SEGMENT_SECONDS = 300   # длинная дорожка пересчитывается по 5 минут (экономия памяти)

SILENCE_PEAK = 0.001         # тише этого уровня считаем, что звук вообще не записался
TEMP_PREFIX = "lesson-transcriber-"   # префикс временной папки

WINDOW_TITLE = "Транскрипция урока"

STATUS_READY = "Готов к записи"
STATUS_RECORDING = "Идёт запись..."
STATUS_PROCESSING = "Обработка транскрипции..."

# Подсказки, которые показываются, если что-то не работает.
ADVICE_SYSTEM_SOUND = (
    "Что можно попробовать:\n"
    "1. Убедитесь, что звук урока идёт в устройство по умолчанию — то, которое выбрано\n"
    "   в Windows: «Параметры → Система → Звук → Устройство вывода».\n"
    "2. Если урок слушается в Bluetooth-наушниках, попробуйте проводные наушники\n"
    "   или встроенные динамики: у некоторых Bluetooth-устройств захват невозможен.\n"
    "3. Закройте программы, которые могут «занимать» звуковую карту целиком.\n"
    "4. Если не помогает — установите виртуальный кабель VB-Cable (ссылка в README.md).\n"
    "   Иногда это единственный выход, если встроенный режим Windows\n"
    "   «WASAPI Loopback» на вашей звуковой карте не работает."
)

ADVICE_MICROPHONE = (
    "Что можно попробовать:\n"
    "1. Подключите микрофон (или гарнитуру) к компьютеру.\n"
    "2. В Windows проверьте: «Параметры → Конфиденциальность → Микрофон» —\n"
    "   доступ к микрофону должен быть разрешён.\n"
    "3. Проверьте «Параметры → Система → Звук → Устройство ввода» —\n"
    "   выбран ли нужный микрофон, и двигается ли индикатор громкости, когда вы говорите.\n"
    "4. Перезапустите программу после подключения микрофона."
)


def log(message: str) -> None:
    """Печатает сообщение в консоль (если она есть) — это помогает понять, что происходит."""
    try:
        print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)
    except Exception:  # консоли может не быть вовсе (запуск через pythonw)
        pass


# ============================================================
#  2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================


class TranscriptionError(Exception):
    """Ошибка распознавания, текст которой можно показать пользователю как есть."""


def format_hms(seconds: float) -> str:
    """Превращает количество секунд в строку вида 01:23:45."""
    seconds = int(max(0, seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def short_error(exc: BaseException) -> str:
    """Возвращает короткое описание ошибки (без внутренних путей Python)."""
    text = str(exc).strip() or exc.__class__.__name__
    text = re.sub(r"\s+", " ", text)
    return text[:300]


# ---------- путь к рабочему столу ----------

class _GUID(ctypes.Structure):
    """Служебная структура Windows для функции SHGetKnownFolderPath."""

    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _windows_desktop_path() -> str | None:
    """
    Спрашивает у Windows настоящий путь к рабочему столу.

    Такой способ правильно работает и когда рабочий стол перенесён в OneDrive.
    """
    if sys.platform != "win32":
        return None
    try:
        folder_id = _GUID(
            0xB4BFCC3A, 0xDB2C, 0x424C,
            (ctypes.c_ubyte * 8)(0xB0, 0x29, 0x7F, 0xE9, 0x9A, 0x87, 0xC6, 0x41),
        )
        path_ptr = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(folder_id), 0, None, ctypes.byref(path_ptr)
        )
        if result != 0 or not path_ptr.value:
            return None
        path = str(path_ptr.value)
        ctypes.windll.ole32.CoTaskMemFree(path_ptr)
        return path
    except Exception:
        return None


def get_desktop_dir() -> str:
    """
    Возвращает папку, в которую складывать готовый файл.

    Сначала пробуем настоящий рабочий стол Windows, затем OneDrive\\Desktop,
    затем обычную папку Desktop в профиле пользователя.
    """
    override = os.environ.get("LESSON_TRANSCRIBER_OUTPUT_DIR")
    if override:
        os.makedirs(override, exist_ok=True)
        return override

    home = os.path.expanduser("~")
    candidates = [
        _windows_desktop_path(),
        os.path.join(os.environ.get("OneDrive", ""), "Desktop") if os.environ.get("OneDrive") else None,
        os.path.join(home, "Desktop"),
        os.path.join(home, "Рабочий стол"),
        home,
    ]
    for candidate in candidates:
        if candidate and os.path.isdir(candidate) and os.access(candidate, os.W_OK):
            return candidate

    # Если папки нет — попробуем создать «Desktop» и в крайнем случае используем домашнюю папку.
    fallback = os.path.join(home, "Desktop")
    try:
        os.makedirs(fallback, exist_ok=True)
        return fallback
    except Exception:
        return home


def unique_path(folder: str, base_name: str, extension: str) -> str:
    """Придумывает имя файла, которое ещё не занято (добавляет «(2)», «(3)» и т. д.)."""
    path = os.path.join(folder, base_name + extension)
    counter = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{base_name} ({counter}){extension}")
        counter += 1
    return path


# ---------- работа с аудио ----------


def to_mono_int16(block) -> np.ndarray:
    """
    Приводит любой полученный от устройства блок звука к моно-формату int16.

    Sounddevice и PyAudio отдают целые числа (int16), а библиотека soundcard —
    дробные числа от -1 до 1. Приводим всё к одному виду.
    """
    arr = np.asarray(block)
    if np.issubdtype(arr.dtype, np.floating):
        if arr.ndim == 2 and arr.shape[1] > 1:      # усредняем каналы (стерео → моно)
            arr = arr.mean(axis=1, dtype=np.float32)
        arr = np.reshape(arr, -1)
        return np.clip(arr * 32767.0, -32768.0, 32767.0).astype(np.int16)

    if arr.ndim == 2 and arr.shape[1] > 1:
        arr = arr.mean(axis=1, dtype=np.float32)
        return np.clip(np.rint(arr), -32768, 32767).astype(np.int16)
    return np.reshape(arr, -1).astype(np.int16, copy=False)


def float_to_int16(data: np.ndarray) -> np.ndarray:
    """Переводит массив из диапазона -1..1 в 16-битные целые числа."""
    return np.clip(np.asarray(data, dtype=np.float32) * 32767.0, -32768.0, 32767.0).astype(np.int16)


def resample_audio(data: np.ndarray, source_rate: int, target_rate: int) -> np.ndarray:
    """
    Пересчитывает звук на другую частоту дискретизации.

    Нужно редко: обычно и системный звук, и микрофон пишутся сразу в 16 кГц.
    Но если какое-то устройство умеет работать только на 48 кГц — пересчитываем сами.
    """
    data = np.reshape(np.asarray(data, dtype=np.float32), -1)
    if source_rate == target_rate or data.size == 0:
        return data

    count = int(round(data.size * target_rate / source_rate))
    if count <= 0:
        return np.zeros(0, dtype=np.float32)

    # Если библиотека scipy есть — используем её качественный пересчёт (быстро и точно).
    try:  # pragma: no cover — зависит от того, установлена ли scipy
        from scipy import signal  # type: ignore

        return np.asarray(signal.resample_poly(data, target_rate, source_rate), dtype=np.float32)
    except Exception:
        pass

    # Запасной вариант без дополнительных библиотек: простое сглаживание
    # (чтобы не было «звона» при понижении частоты) + линейная интерполяция.
    if target_rate < source_rate:
        step = int(round(source_rate / target_rate))
        if 2 <= step <= 8:
            kernel = np.ones(step, dtype=np.float32) / step
            data = np.convolve(data, kernel, mode="same").astype(np.float32)

    old_grid = np.arange(data.size, dtype=np.float64)
    new_grid = np.linspace(0.0, data.size - 1, count, dtype=np.float64)
    return np.interp(new_grid, old_grid, data).astype(np.float32)


def resample_wav(source_path: str, target_path: str, target_rate: int = SAMPLE_RATE) -> None:
    """
    Пересохраняет WAV-файл в нужной частоте дискретизации.

    Файл обрабатывается кусками по 5 минут, чтобы даже очень длинная запись
    (например, двухчасовой урок) не требовала много оперативной памяти.
    """
    info = sf.info(source_path)
    if info.samplerate == target_rate:
        shutil.copyfile(source_path, target_path)
        return

    segment_frames = int(RESAMPLE_SEGMENT_SECONDS * info.samplerate)
    with sf.SoundFile(source_path) as source, sf.SoundFile(
        target_path, mode="w", samplerate=target_rate, channels=1, subtype="PCM_16"
    ) as target:
        while True:
            data = source.read(segment_frames, dtype="float32", always_2d=False)
            if data.size == 0:
                break
            converted = resample_audio(np.reshape(data, -1), info.samplerate, target_rate)
            target.write(float_to_int16(converted))


def _wav_blocks(path: str, block_frames: int):
    """Читает WAV-файл блоками (чтобы не держать весь урок в оперативной памяти)."""
    with sf.SoundFile(path) as handle:
        while True:
            data = handle.read(block_frames, dtype="float32", always_2d=False)
            if data.size == 0:
                break
            yield np.reshape(data, -1)


def _summed_blocks(paths: list[str], total_frames: int, block_frames: int):
    """Отдаёт по очереди блоки, в которых дорожки уже сложены друг с другом."""
    readers = [_wav_blocks(path, block_frames) for path in paths]
    position = 0
    while position < total_frames:
        want = min(block_frames, total_frames - position)
        parts = []
        for reader in readers:
            try:
                block = next(reader)
            except StopIteration:
                block = np.zeros(0, dtype=np.float32)
            block = np.reshape(block, -1)
            if block.size > want:
                block = block[:want]
            elif block.size < want:                     # дорожка закончилась раньше — добиваем тишиной
                block = np.pad(block, (0, want - block.size))
            parts.append(block)
        position += want
        yield parts[0] + parts[1] if len(parts) == 2 else parts[0]


def mix_and_normalize(paths: list[str], output_path: str, block_frames: int = MIX_BLOCK_FRAMES):
    """
    Складывает дорожки (системный звук + микрофон), нормализует громкость
    и сохраняет результат одним моно-WAV 16 кГц / 16 бит.

    Возвращает (пиковая громкость, длительность в секундах).
    Сначала проходим по файлам и находим самый громкий момент, потом записываем
    результат с нужным усилением — поэтому количество проходов равно двум.
    """
    infos = [sf.info(path) for path in paths]
    lengths = [info.frames for info in infos]

    # Обычный случай — обе дорожки одинаковой длины, обрезаем по самой короткой (как в ТЗ).
    # Но если одна дорожка оборвалась сильно раньше (например, устройство отключили),
    # лучше добить её тишиной, чем потерять вторую половину урока.
    if len(lengths) > 1 and (max(lengths) - min(lengths)) > 2 * SAMPLE_RATE:
        total_frames = max(lengths)
        log("Внимание: дорожки разной длины — короткая дополнена тишиной.")
    else:
        total_frames = min(lengths)

    if total_frames <= 0:
        return 0.0, 0.0

    # 1-й проход: ищем пик, чтобы понять нужное усиление.
    peak = 0.0
    for block in _summed_blocks(paths, total_frames, block_frames):
        if block.size:
            peak = max(peak, float(np.max(np.abs(block))))

    if peak <= 0:
        gain = 1.0
    elif peak > 0.95:                     # громко — только уменьшаем, чтобы не было хрипа
        gain = 0.95 / peak
    elif peak < 0.02:                     # очень тихо — аккуратно поднимаем громкость
        gain = min(0.30 / peak, 50.0)
    else:
        gain = 1.0

    # 2-й проход: записываем итоговый файл.
    with sf.SoundFile(
        output_path, mode="w", samplerate=SAMPLE_RATE, channels=1, subtype="PCM_16"
    ) as out:
        for block in _summed_blocks(paths, total_frames, block_frames):
            out.write(float_to_int16(block * gain))

    return peak, total_frames / SAMPLE_RATE


# ---------- текст ----------

def join_segments(segments: list[tuple[float, float, str]], paragraph_gap: float = 2.5) -> str:
    """
    Собирает распознанные фразы в читаемый текст.

    Фразы, между которыми пауза больше 2,5 секунды, начинают новый абзац.
    Таймкоды и служебные пометки в файл не попадают.
    """
    paragraphs: list[str] = []
    current: list[str] = []
    previous_end: float | None = None

    for start, end, text in segments:
        text = (text or "").strip()
        if not text:
            continue
        if previous_end is not None and (start - previous_end) > paragraph_gap and current:
            paragraphs.append(" ".join(current))
            current = []
        current.append(text)
        previous_end = max(previous_end or 0.0, end)

    if current:
        paragraphs.append(" ".join(current))

    text = "\n\n".join(paragraphs)
    text = text.replace("♪", " ")                       # Whisper иногда помечает музыку нотками
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)         # убираем пробел перед запятой
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def cleanup_old_temp_dirs() -> None:
    """Удаляет временные папки прошлых запусков, которые остались после сбоя."""
    temp_root = tempfile.gettempdir()
    try:
        names = os.listdir(temp_root)
    except Exception:
        return
    for name in names:
        if not name.startswith(TEMP_PREFIX):
            continue
        path = os.path.join(temp_root, name)
        try:
            if os.path.isdir(path) and time.time() - os.path.getmtime(path) > 24 * 3600:
                shutil.rmtree(path, ignore_errors=True)
        except Exception:
            continue


# ============================================================
#  3. ЗАХВАТ ЗВУКА
# ============================================================


def init_com_for_thread() -> bool:
    """
    Инициализирует COM в текущем потоке.

    Это нужно на Windows: библиотеки захвата звука обращаются к служебным
    компонентам Windows и отказываются работать в потоке, где COM не
    подготовлен. Возвращает True, если инициализацию выполнили именно мы —
    тогда её нужно отменить в конце работы потока.
    """
    if sys.platform != "win32":
        return False
    try:
        result = ctypes.windll.ole32.CoInitializeEx(None, 0x2)  # COINIT_APARTMENTTHREADED
        return result in (0, 1)  # S_OK или S_FALSE
    except Exception:
        return False


def uninit_com_for_thread(initialized: bool) -> None:
    """Завершает работу с COM в потоке (парная функция к init_com_for_thread)."""
    if not initialized or sys.platform != "win32":
        return
    try:
        ctypes.windll.ole32.CoUninitialize()
    except Exception:
        pass


class BaseRecorder(threading.Thread):
    """
    Общая часть для записи звука.

    Каждый источник пишет свой поток в отдельном потоке и сразу сохраняет его
    в WAV-файл — поэтому даже долгая запись не занимает много оперативной памяти.
    """

    display_name = "Источник звука"

    def __init__(self, wav_path: str):
        super().__init__(daemon=True, name=f"rec-{os.path.basename(wav_path)}")
        self.wav_path = wav_path
        self.sample_rate = SAMPLE_RATE
        self.frames_written = 0
        self.error: BaseException | None = None
        self.opened = threading.Event()      # взводится, когда устройство открыто (или не открылось)
        self.stop_event = threading.Event()  # взводится, когда пользователь нажал «Остановить»
        self._device_name = ""

    # --- методы, которые реализуют наследники ---
    def open(self) -> None:
        """Открывает устройство. Должен установить self.sample_rate и self._device_name."""
        raise NotImplementedError

    def read(self) -> np.ndarray:
        """Читает очередной блок звука (моно, int16). Вызывается в цикле."""
        raise NotImplementedError

    def close(self) -> None:
        """Освобождает устройство."""
        return None

    # --- общая логика ---
    def run(self) -> None:
        com_initialized = init_com_for_thread()
        try:
            try:
                self.open()
            except BaseException as exc:  # noqa: BLE001 — показываем пользователю любую причину
                self.error = exc
                log(f"{self.display_name}: не удалось открыть устройство — {short_error(exc)}")
            else:
                self.opened.set()
                log(f"{self.display_name}: пишем с «{self._device_name}», "
                    f"частота {self.sample_rate} Гц")
                try:
                    with sf.SoundFile(
                        self.wav_path, mode="w", samplerate=self.sample_rate,
                        channels=1, subtype="PCM_16",
                    ) as out:
                        while not self.stop_event.is_set():
                            block = self.read()
                            if block is None or len(block) == 0:
                                continue
                            out.write(block)
                            self.frames_written += len(block)
                except BaseException as exc:  # noqa: BLE001
                    self.error = exc
                    log(f"{self.display_name}: запись прервана — {short_error(exc)}")
                finally:
                    try:
                        self.close()
                    except Exception:
                        pass
            self.opened.set()  # при ошибке тот, кто ждёт, тоже должен проснуться
        finally:
            uninit_com_for_thread(com_initialized)

    def duration(self) -> float:
        """Сколько секунд записано."""
        return self.frames_written / self.sample_rate if self.sample_rate else 0.0


class MicRecorder(BaseRecorder):
    """Запись с микрофона (библиотека sounddevice)."""

    display_name = "Микрофон"

    def __init__(self, wav_path: str):
        super().__init__(wav_path)
        self._stream = None

    def open(self) -> None:
        try:
            import sounddevice as sd
        except ImportError as exc:
            raise RuntimeError(
                "Не установлена библиотека sounddevice (нужна для записи с микрофона).\n"
                "Выполните команду: pip install -r requirements.txt"
            ) from exc

        device = find_input_device(sd)
        if device is None:
            raise RuntimeError("В системе не найден ни один микрофон.")

        # Пытаемся открыть микрофон сразу на 16 кГц — столько нужно для
        # распознавания. Если устройство не умеет, берём его частоту
        # и позже аккуратно пересчитаем запись.
        native_rate = int(round(float(device.get("default_samplerate") or SAMPLE_RATE)))
        max_channels = max(1, int(device.get("max_input_channels") or 1))
        candidates = [(SAMPLE_RATE, 1), (native_rate, 1), (native_rate, min(2, max_channels))]
        candidates = list(dict.fromkeys(candidates))  # убираем повторы, сохраняя порядок

        last_error: BaseException | None = None
        for rate, channels in candidates:
            stream = None
            try:
                stream = sd.InputStream(
                    samplerate=rate,
                    channels=channels,
                    dtype="int16",
                    device=device["index"],
                    blocksize=max(64, int(rate * RECORD_BLOCK_SECONDS)),
                )
                stream.start()
            except Exception as exc:  # noqa: BLE001 — пробуем следующий вариант
                last_error = exc
                if stream is not None:
                    try:
                        stream.close()
                    except Exception:
                        pass
                continue
            self._stream = stream
            self.sample_rate = rate
            self._device_name = str(device.get("name", "микрофон"))
            return

        raise RuntimeError(
            "Не удалось включить микрофон"
            + (f": {short_error(last_error)}" if last_error else "")
        )

    def read(self) -> np.ndarray:
        data, overflowed = self._stream.read(max(64, int(self.sample_rate * RECORD_BLOCK_SECONDS)))
        if overflowed:
            log("Микрофон: переполнение буфера (звук всё равно записывается).")
        return to_mono_int16(data)

    def close(self) -> None:
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None


class SystemAudioRecorder(BaseRecorder):
    """
    Запись системного звука (то, что слышно в колонках или наушниках).

    Пробуем несколько способов по очереди: если один не сработал —
    автоматически переходим к следующему. Драйверы устанавливать не нужно:
    используется встроенный в Windows режим WASAPI Loopback.
    """

    display_name = "Системный звук"

    def __init__(self, wav_path: str):
        super().__init__(wav_path)
        self._backend = None

    def open(self) -> None:
        problems = []
        for backend_cls in (PyAudioWPatchLoopback, SoundCardLoopback, SoundDeviceLoopback):
            backend = backend_cls()
            try:
                backend.open()
            except Exception as exc:  # noqa: BLE001 — пробуем следующий способ
                problems.append(f"— {backend_cls.label}: {short_error(exc)}")
                try:
                    backend.close()
                except Exception:
                    pass
                continue
            self._backend = backend
            self.sample_rate = backend.sample_rate
            self._device_name = f"{backend.device_name} [{backend_cls.label}]"
            return

        raise RuntimeError(
            "Не удалось захватить системный звук.\n" + "\n".join(problems)
        )

    def read(self) -> np.ndarray:
        return self._backend.read()

    def close(self) -> None:
        if self._backend is not None:
            try:
                self._backend.close()
            except Exception:
                pass
            self._backend = None


# ---------- три способа захватить системный звук ----------


class PyAudioWPatchLoopback:
    """Способ 1: библиотека PyAudioWPatch (самый надёжный вариант на Windows)."""

    label = "PyAudioWPatch"

    def __init__(self):
        self._pa = None
        self._stream = None
        self.device_name = ""
        self.sample_rate = SAMPLE_RATE
        self._channels = 1

    @staticmethod
    def _import():
        try:
            import pyaudiowpatch as pyaudio  # noqa: N813
        except ImportError as exc:
            raise RuntimeError(
                "не установлена библиотека PyAudioWPatch (pip install -r requirements.txt)"
            ) from exc
        return pyaudio

    @classmethod
    def probe(cls) -> str:
        """Проверяет, что устройство записи системного звука есть в системе."""
        backend = cls()
        try:
            backend._pa = backend._import().PyAudio()
            return str(backend._find_device().get("name", "системный звук"))
        finally:
            backend.close()

    def open(self) -> None:
        pyaudio = self._import()
        self._pa = pyaudio.PyAudio()
        try:
            device = self._find_device()
            index = int(device["index"])
            max_channels = max(1, min(2, int(device.get("maxInputChannels") or 1)))
            native_rate = int(round(float(device.get("defaultSampleRate") or SAMPLE_RATE)))
            rates = [SAMPLE_RATE] if native_rate == SAMPLE_RATE else [SAMPLE_RATE, native_rate]

            last_error: BaseException | None = None
            for rate in rates:
                for channels in sorted({max_channels, 1}, reverse=True):
                    try:
                        self._stream = self._pa.open(
                            format=pyaudio.paInt16,
                            channels=channels,
                            rate=rate,
                            input=True,
                            input_device_index=index,
                            frames_per_buffer=max(256, int(rate * RECORD_BLOCK_SECONDS)),
                        )
                    except Exception as exc:  # noqa: BLE001 — пробуем следующий вариант
                        last_error = exc
                        self._stream = None
                        continue
                    self.sample_rate = rate
                    self._channels = channels
                    self.device_name = str(device.get("name", "системный звук"))
                    return

            raise RuntimeError(
                "не удалось открыть поток системного звука"
                + (f": {short_error(last_error)}" if last_error else "")
            )
        except Exception:
            self.close()
            raise

    def _find_device(self):
        """
        Ищет устройство «динамики в режиме записи» (WASAPI Loopback).

        Сначала спрашиваем у Windows устройство по умолчанию, затем ищем
        его же среди loopback-устройств, затем берём любое подходящее.
        """
        candidates = []
        try:
            candidates.append(self._pa.get_default_wasapi_loopback())
        except Exception:
            pass

        try:
            pyaudio = self._import()
            wasapi_info = self._pa.get_host_api_info_by_type(pyaudio.paWASAPI)
            default_output = int(wasapi_info.get("defaultOutputDevice", -1))
            if default_output >= 0:
                output_name = str(self._pa.get_device_info_by_index(default_output).get("name", ""))
                for loopback in self._pa.get_loopback_device_info_generator():
                    if output_name and output_name in str(loopback.get("name", "")):
                        candidates.append(loopback)
                        break
        except Exception:
            pass

        try:
            for loopback in self._pa.get_loopback_device_info_generator():
                candidates.append(loopback)
        except Exception:
            pass

        for device in candidates:
            if device and int(device.get("maxInputChannels", 0) or 0) > 0:
                return device
        raise RuntimeError("в системе нет устройства захвата системного звука (WASAPI Loopback)")

    def read(self) -> np.ndarray:
        raw = self._stream.read(
            max(256, int(self.sample_rate * RECORD_BLOCK_SECONDS)), exception_on_overflow=False
        )
        data = np.frombuffer(raw, dtype=np.int16)
        if self._channels > 1:
            data = data.reshape(-1, self._channels)
        return to_mono_int16(data)

    def close(self) -> None:
        if self._stream is not None:
            try:
                self._stream.stop_stream()
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        if self._pa is not None:
            try:
                self._pa.terminate()
            except Exception:
                pass
            self._pa = None


class SoundCardLoopback:
    """Способ 2: библиотека soundcard (тоже умеет WASAPI Loopback)."""

    label = "soundcard"

    def __init__(self):
        self._recorder = None
        self._context = None
        self.device_name = ""
        self.sample_rate = SAMPLE_RATE
        self._channels = 2

    @staticmethod
    def _import():
        try:
            import soundcard as sc
        except ImportError as exc:
            raise RuntimeError(
                "не установлена библиотека soundcard (pip install -r requirements.txt)"
            ) from exc
        return sc

    @classmethod
    def _loopback_microphone(cls):
        sc = cls._import()
        speaker = sc.default_speaker()
        if speaker is None:
            raise RuntimeError("в системе не найдено устройство воспроизведения по умолчанию")
        loopback = sc.get_microphone(id=str(speaker.name), include_loopback=True)
        if loopback is None:
            raise RuntimeError("не удалось получить доступ к системному звуку (loopback)")
        return speaker, loopback

    @classmethod
    def probe(cls) -> str:
        speaker, _loopback = cls._loopback_microphone()
        return str(getattr(speaker, "name", "системный звук"))

    def open(self) -> None:
        speaker, loopback = self._loopback_microphone()
        self._channels = max(1, min(2, int(getattr(speaker, "channels", 2) or 2)))
        self._context = loopback.recorder(samplerate=SAMPLE_RATE, channels=self._channels)
        self._recorder = self._context.__enter__()
        self.sample_rate = SAMPLE_RATE
        self.device_name = str(getattr(speaker, "name", "системный звук"))

    def read(self) -> np.ndarray:
        # Просим всегда одинаковый по длине блок и при необходимости дополняем
        # его тишиной: эта библиотека иногда возвращает меньше данных, чем
        # просили, а «сдвиг» звука по времени нам не нужен.
        wanted = max(64, int(self.sample_rate * RECORD_BLOCK_SECONDS))
        data = self._recorder.record(numframes=wanted)
        if data is None or len(data) == 0:
            return np.zeros(wanted, dtype=np.int16)
        block = to_mono_int16(np.asarray(data, dtype=np.float32))
        if block.size < wanted:
            block = np.pad(block, (0, wanted - block.size))
        elif block.size > wanted:
            block = block[:wanted]
        return block

    def close(self) -> None:
        if self._context is not None:
            try:
                self._context.__exit__(None, None, None)
            except Exception:
                pass
            self._context = None
            self._recorder = None


class SoundDeviceLoopback:
    """
    Способ 3: захват системного звука через sounddevice.

    Работает в новых версиях sounddevice (0.5 и выше), которые умеют
    открывать устройство вывода в режиме loopback.
    """

    label = "sounddevice (loopback)"

    def __init__(self):
        self._stream = None
        self.device_name = ""
        self.sample_rate = SAMPLE_RATE
        self._channels = 2

    @staticmethod
    def _check_support():
        """Проверяет, что библиотека sounddevice вообще умеет loopback-запись."""
        try:
            import sounddevice as sd
        except ImportError as exc:
            raise RuntimeError("не установлена библиотека sounddevice") from exc

        if not hasattr(sd, "WasapiSettings"):
            raise RuntimeError("эта версия sounddevice не поддерживает запись системного звука")

        try:  # параметр loopback появился в sounddevice 0.5
            sd.WasapiSettings(loopback=True)
        except TypeError as exc:
            raise RuntimeError("эта версия sounddevice не умеет писать системный звук") from exc
        return sd

    @classmethod
    def probe(cls) -> str:
        """Проверяет доступность записи системного звука, ничего не записывая."""
        sd = cls._check_support()
        output = find_output_device(sd)
        if output is None:
            raise RuntimeError("в системе не найдено устройство воспроизведения")
        return str(output.get("name", "системный звук"))

    def open(self) -> None:
        sd = self._check_support()
        output = find_output_device(sd)
        if output is None:
            raise RuntimeError("в системе не найдено устройство воспроизведения")

        native_rate = int(round(float(output.get("default_samplerate") or SAMPLE_RATE)))
        max_channels = max(1, min(2, int(output.get("max_output_channels") or 2)))
        last_error: BaseException | None = None

        for rate in dict.fromkeys([SAMPLE_RATE, native_rate]):
            for channels in sorted({max_channels, 1}, reverse=True):
                try:
                    settings = sd.WasapiSettings(loopback=True)
                    stream = sd.InputStream(
                        samplerate=rate,
                        channels=channels,
                        dtype="int16",
                        device=output["index"],
                        blocksize=max(64, int(rate * RECORD_BLOCK_SECONDS)),
                        extra_settings=settings,
                    )
                    stream.start()
                except Exception as exc:  # noqa: BLE001 — пробуем следующий вариант
                    last_error = exc
                    continue
                self._stream = stream
                self.sample_rate = rate
                self._channels = channels
                self.device_name = str(output.get("name", "системный звук"))
                return

        raise RuntimeError(
            "не удалось открыть loopback-поток"
            + (f": {short_error(last_error)}" if last_error else "")
        )

    def read(self) -> np.ndarray:
        data, overflowed = self._stream.read(max(64, int(self.sample_rate * RECORD_BLOCK_SECONDS)))
        if overflowed:
            log("Системный звук: переполнение буфера.")
        return to_mono_int16(data)

    def close(self) -> None:
        if self._stream is not None:
            try:
                self._stream.stop()
                self._stream.close()
            except Exception:
                pass
            self._stream = None


# ---------- выбор устройств и проверка перед записью ----------


def ensure_device_index(device, sd):
    """
    Убеждаемся, что в описании устройства есть его номер.

    Библиотека sounddevice не всегда добавляет поле «index», а оно нужно,
    чтобы открыть именно это устройство. Если описания нет — вернём None.
    """
    if not device:
        return None
    device = dict(device)
    if device.get("index") is not None:
        return device
    try:
        for index, candidate in enumerate(sd.query_devices()):
            if candidate.get("name") == device.get("name"):
                device["index"] = index
                break
    except Exception:
        pass
    return device if device.get("index") is not None else None


def is_input_device(device) -> bool:
    """Годится ли устройство для записи."""
    return bool(device) and int(device.get("max_input_channels", 0) or 0) > 0


def is_output_device(device) -> bool:
    """Годится ли устройство для воспроизведения."""
    return bool(device) and int(device.get("max_output_channels", 0) or 0) > 0


def find_input_device(sd):
    """Выбирает микрофон по умолчанию (по возможности через WASAPI)."""
    # 1. Устройство ввода по умолчанию в WASAPI — оно надёжнее всего.
    try:
        for host in sd.query_hostapis():
            if "WASAPI" in str(host.get("name", "")).upper():
                index = int(host.get("default_input_device", -1))
                if index >= 0:
                    device = ensure_device_index(sd.query_devices(index), sd)
                    if is_input_device(device):
                        return device
    except Exception:
        pass

    # 2. Микрофон, выбранный в Windows по умолчанию.
    try:
        device = ensure_device_index(sd.query_devices(kind="input"), sd)
        if is_input_device(device):
            return device
    except Exception:
        pass

    # 3. Любое устройство, которое умеет записывать звук.
    try:
        for index, device in enumerate(sd.query_devices()):
            if int(device.get("max_input_channels", 0) or 0) > 0:
                device = dict(device)
                device["index"] = index
                return device
    except Exception:
        pass
    return None


def find_output_device(sd):
    """Выбирает устройство воспроизведения по умолчанию."""
    try:
        for host in sd.query_hostapis():
            if "WASAPI" in str(host.get("name", "")).upper():
                index = int(host.get("default_output_device", -1))
                if index >= 0:
                    device = ensure_device_index(sd.query_devices(index), sd)
                    if is_output_device(device):
                        return device
    except Exception:
        pass

    try:
        device = ensure_device_index(sd.query_devices(kind="output"), sd)
        if is_output_device(device):
            return device
    except Exception:
        pass

    try:
        for index, device in enumerate(sd.query_devices()):
            if int(device.get("max_output_channels", 0) or 0) > 0:
                device = dict(device)
                device["index"] = index
                return device
    except Exception:
        pass
    return None


def check_microphone() -> tuple[bool, str]:
    """
    Проверяет, есть ли микрофон (до начала записи).

    Возвращает пару (всё хорошо, пояснение на русском).
    """
    try:
        import sounddevice as sd
    except ImportError:
        return False, (
            "Не установлена библиотека sounddevice, без неё запись микрофона невозможна.\n\n"
            "Выполните команду: pip install -r requirements.txt"
        )

    try:
        device = find_input_device(sd)
    except Exception as exc:  # noqa: BLE001
        return False, f"Не удалось проверить микрофоны: {short_error(exc)}"

    if device is None:
        return False, "Микрофон не найден.\n\n" + ADVICE_MICROPHONE
    return True, str(device.get("name", "микрофон"))


def check_system_audio() -> tuple[bool, str]:
    """Проверяет, можно ли записывать системный звук (WASAPI Loopback)."""
    problems = []
    for backend in (PyAudioWPatchLoopback, SoundCardLoopback, SoundDeviceLoopback):
        try:
            name = backend.probe()
        except Exception as exc:  # noqa: BLE001 — пробуем следующий способ
            problems.append(f"— {backend.label}: {short_error(exc)}")
            continue
        return True, name

    message = (
        "Не удалось начать захват системного звука (звука Zoom или браузера).\n\n"
        + ADVICE_SYSTEM_SOUND
    )
    if problems:
        message += "\n\nПодробности для специалиста:\n" + "\n".join(problems)
    return False, message


class RecordingSession:
    """Две дорожки одной записи: системный звук и микрофон."""

    def __init__(self, temp_dir: str, use_system: bool, use_mic: bool):
        self.temp_dir = temp_dir
        self.started_at = datetime.now()
        self.sources: list[BaseRecorder] = []
        # Замок нужен потому, что кнопки «Остановить» и закрытие окна могут
        # сработать в тот же миг, когда запись ещё только запускается.
        self._lock = threading.Lock()
        if use_system:
            self.sources.append(SystemAudioRecorder(os.path.join(temp_dir, "system.wav")))
        if use_mic:
            self.sources.append(MicRecorder(os.path.join(temp_dir, "microphone.wav")))

    def start(self) -> None:
        """Запускает потоки записи и ждёт, пока устройства откроются."""
        with self._lock:
            for source in self.sources:
                if not source.stop_event.is_set() and source.ident is None:
                    source.start()
        for source in self.sources:
            source.opened.wait(timeout=20)

    def stop(self) -> None:
        """Просит потоки остановиться (команда выполняется быстро)."""
        with self._lock:
            for source in self.sources:
                # Флаг ставим даже не запущенному потоку: если он всё-таки
                # начнётся, то сразу же завершится, ничего не записывая.
                source.stop_event.set()

    def join(self, timeout: float = 30.0) -> None:
        """Дожидается, пока потоки действительно закончат запись."""
        with self._lock:
            started = [source for source in self.sources if source.ident is not None]
        for source in started:
            source.join(timeout)

    def duration(self) -> float:
        """Длительность записи в секундах."""
        return max((source.duration() for source in self.sources), default=0.0)

    def saved_files(self) -> list[str]:
        """Файлы, которые удалось записать."""
        return [
            source.wav_path
            for source in self.sources
            if os.path.exists(source.wav_path) and os.path.getsize(source.wav_path) > 44
        ]

    def failed_sources(self) -> list[str]:
        """Понятные сообщения о дорожках, которые не записались."""
        messages = []
        for source in self.sources:
            if source.error is not None:
                messages.append(f"{source.display_name}: {short_error(source.error)}")
        return messages


# ============================================================
#  4. РАСПОЗНАВАНИЕ РЕЧИ (faster-whisper)
# ============================================================

_model = None
_model_lock = threading.Lock()


def friendly_model_error(exc: BaseException) -> str:
    """Переводит техническую ошибку загрузки модели на понятный русский язык."""
    text = str(exc)
    lowered = text.lower()

    if "no module named" in lowered and "faster" in lowered:
        return (
            "Не установлена библиотека faster-whisper.\n"
            "Выполните команду: pip install -r requirements.txt"
        )
    if any(word in lowered for word in
           ("localentrynotfound", "local_files_only", "cannot find the requested files")):
        cache_dir = os.path.join(os.path.expanduser("~"), ".cache", "huggingface", "hub")
        return (
            "Скачанная ранее модель распознавания оказалась неполной.\n\n"
            "Подключитесь к интернету и нажмите «Начать запись» ещё раз — модель скачается заново.\n"
            "Если не поможет, удалите папку с моделью и попробуйте снова:\n"
            f"{cache_dir}"
        )
    if any(word in lowered for word in
           ("connection", "connect", "timed out", "timeout", "resolve", "proxy",
            "ssl", "certificate", "offline", "network", "unreachable")):
        return (
            "Не удалось скачать модель распознавания — похоже, нет доступа в интернет.\n\n"
            "Интернет нужен только в первый раз, чтобы скачать модель (около 500 МБ).\n"
            "Проверьте подключение и нажмите «Начать запись» ещё раз."
        )
    if any(word in lowered for word in ("no space", "disk full", "enospc")):
        return (
            "На диске закончилось место — модель распознавания не удалось сохранить.\n"
            "Освободите несколько гигабайт и попробуйте снова."
        )
    if any(word in lowered for word in ("memory", "bad_alloc", "out of memory")):
        return (
            "Компьютеру не хватило оперативной памяти для модели распознавания.\n\n"
            "Закройте лишние программы и попробуйте снова. Если не поможет,\n"
            "можно уменьшить модель: откройте main.py, найдите строку MODEL_SIZE = \"small\"\n"
            "и замените \"small\" на \"base\"."
        )
    return f"Не удалось подготовить модель распознавания:\n{short_error(exc)}"


def model_repo_id() -> str:
    """Возвращает имя модели в интернете (например, Systran/faster-whisper-small)."""
    try:
        from faster_whisper.utils import _MODELS  # type: ignore

        return str(_MODELS.get(MODEL_SIZE, f"Systran/faster-whisper-{MODEL_SIZE}"))
    except Exception:
        return f"Systran/faster-whisper-{MODEL_SIZE}"


def is_model_downloaded() -> bool:
    """
    Проверяет, скачана ли модель раньше.

    Нужно только для того, чтобы показать правильную подсказку
    («Скачивание модели...» или «Загрузка модели...»).
    """
    repo_id = model_repo_id()
    try:
        from huggingface_hub import try_to_load_from_cache  # type: ignore

        result = try_to_load_from_cache(repo_id, "model.bin")
        if isinstance(result, str):
            return True
    except Exception:
        pass

    # Запасная проверка: ищем папку модели в кэше Hugging Face вручную.
    try:
        cache_root = os.environ.get("HF_HUB_CACHE") or os.environ.get("HF_HOME")
        if not cache_root:
            cache_root = os.path.join(os.path.expanduser("~"), ".cache", "huggingface")
        if os.path.basename(cache_root) != "hub":
            cache_root = os.path.join(cache_root, "hub")
        folder = "models--" + repo_id.replace("/", "--")
        model_dir = os.path.join(cache_root, folder, "snapshots")
        for name in os.listdir(model_dir):
            if os.path.exists(os.path.join(model_dir, name, "model.bin")):
                return True
    except Exception:
        pass
    return False


def load_model(status_callback) -> object:
    """
    Загружает модель Whisper. Делается ЛЕНИВО — только при первой транскрипции,
    чтобы программа запускалась мгновенно.

    Второй и последующие запуски используют уже скачанную модель: повторно
    ничего не скачивается.
    """
    global _model
    with _model_lock:
        if _model is not None:
            return _model

        if is_model_downloaded():
            status_callback("Загрузка модели...")
        else:
            status_callback("Скачивание модели (около 500 МБ, только в первый раз)...")

        try:
            from faster_whisper import WhisperModel  # тяжёлый импорт, поэтому внутри функции

            parameters = {
                "device": "cpu",
                "compute_type": COMPUTE_TYPE,
                "cpu_threads": os.cpu_count() or 4,   # используем все ядра процессора
            }
            # Если модель уже скачана, разрешаем работать без интернета.
            # (параметр есть не во всех версиях faster-whisper, поэтому проверяем)
            try:
                import inspect

                if "local_files_only" in inspect.signature(WhisperModel.__init__).parameters:
                    parameters["local_files_only"] = is_model_downloaded()
            except Exception:
                pass

            log(f"Загружаем модель {MODEL_SIZE} ({COMPUTE_TYPE})...")
            _model = WhisperModel(MODEL_SIZE, **parameters)
            log("Модель готова.")
        except Exception as exc:  # noqa: BLE001 — показываем понятный текст
            raise TranscriptionError(friendly_model_error(exc)) from exc

        return _model


def plan_chunks(wav_path: str, duration: float) -> list[tuple[float, float]]:
    """
    Разбивает длинную запись на куски примерно по 5 минут.

    Границы стараемся ставить в паузах между словами, чтобы не обрезать фразу.
    Так даже трёхчасовой урок обрабатывается надёжно и с небольшим расходом памяти.
    """
    if duration <= TRANSCRIBE_CHUNK_SECONDS * 1.2:
        return [(0.0, float(duration))]

    cuts: list[float] = []
    position = float(TRANSCRIBE_CHUNK_SECONDS)
    while position < duration - TRANSCRIBE_CHUNK_SECONDS * 0.5:
        cut = find_quiet_moment(wav_path, position)
        if cuts and cut - cuts[-1] < TRANSCRIBE_CHUNK_SECONDS * 0.5:
            cut = cuts[-1] + TRANSCRIBE_CHUNK_SECONDS * 0.5  # не делаем слишком коротких кусков
        cut = min(cut, float(duration))
        cuts.append(cut)
        position += TRANSCRIBE_CHUNK_SECONDS

    borders = [0.0] + cuts + [float(duration)]
    return [(borders[i], borders[i + 1]) for i in range(len(borders) - 1) if borders[i + 1] > borders[i]]


def find_quiet_moment(wav_path: str, position: float) -> float:
    """Ищет самый тихий момент рядом с заданным временем (там удобно разрезать запись)."""
    window = 0.25
    step = 0.05
    start_frame = max(0, int((position - BOUNDARY_SEARCH_SECONDS) * SAMPLE_RATE))
    frames = int(BOUNDARY_SEARCH_SECONDS * 2 * SAMPLE_RATE)
    try:
        data, _ = sf.read(wav_path, start=start_frame, frames=frames,
                          dtype="float32", always_2d=False)
    except Exception:
        return position

    data = np.reshape(np.asarray(data, dtype=np.float32), -1)
    window_frames = max(1, int(window * SAMPLE_RATE))
    step_frames = max(1, int(step * SAMPLE_RATE))
    if data.size < window_frames * 3:
        return position

    count = (data.size - window_frames) // step_frames + 1
    starts = np.arange(count) * step_frames
    windows = data[starts[:, None] + np.arange(window_frames)[None, :]]
    energies = np.einsum("ij,ij->i", windows, windows)
    best = int(starts[int(np.argmin(energies))]) + window_frames // 2
    return (start_frame + best) / SAMPLE_RATE


def read_audio_segment(wav_path: str, start: float, end: float) -> np.ndarray:
    """Читает кусок записи как массив float32 в диапазоне -1..1 (формат для Whisper)."""
    start_frame = max(0, int(start * SAMPLE_RATE))
    frames = max(1, int((end - start) * SAMPLE_RATE))
    data, _ = sf.read(wav_path, start=start_frame, frames=frames,
                      dtype="float32", always_2d=False)
    return np.reshape(np.asarray(data, dtype=np.float32), -1)


def transcribe_file(wav_path: str, status_callback, progress_callback) -> str:
    """
    Распознаёт речь в WAV-файле и возвращает готовый текст.

    Файл обрабатывается кусками по 5 минут: faster-whisper и сам умеет работать
    с длинными файлами, но так мы контролируем память и показываем прогресс.
    """
    try:
        duration = sf.info(wav_path).frames / SAMPLE_RATE
    except Exception as exc:  # noqa: BLE001
        raise TranscriptionError(f"Не удалось прочитать запись: {short_error(exc)}") from exc

    if duration <= 0.5:
        raise TranscriptionError("Запись получилась слишком короткой — распознавать нечего.")

    model = load_model(status_callback)

    status_callback(STATUS_PROCESSING)
    progress_callback(0.0)

    chunks = plan_chunks(wav_path, duration)
    log(f"Распознаём {format_hms(duration)} в {len(chunks)} куске(ах).")

    segments: list[tuple[float, float, str]] = []
    for index, (start, end) in enumerate(chunks, start=1):
        audio = read_audio_segment(wav_path, start, end)
        if audio.size == 0:
            continue
        try:
            results, _info = model.transcribe(
                audio,
                language=LANGUAGE,
                task="transcribe",
                beam_size=BEAM_SIZE,
                vad_filter=True,                       # пропускаем тишину
                condition_on_previous_text=False,      # без этого возможны «зацикливания»
                vad_parameters={"min_silence_duration_ms": VAD_MIN_SILENCE_MS},
            )
            # results — генератор: текст появляется по мере распознавания,
            # поэтому прогресс можно показывать «живьём».
            for segment in results:
                text = (getattr(segment, "text", "") or "").strip()
                if text:
                    segments.append((start + float(segment.start), start + float(segment.end), text))
                progress_callback(min(0.999, (start + float(segment.end)) / duration))
            # Даже если в куске была одна тишина, показываем, что он обработан.
            progress_callback(min(0.999, end / duration))
        except TranscriptionError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise TranscriptionError(
                f"Не удалось распознать часть записи (кусок {index} из {len(chunks)}):\n"
                f"{short_error(exc)}"
            ) from exc

    progress_callback(1.0)
    text = join_segments(segments)
    if not text:
        raise TranscriptionError(
            "Речь в записи не распознана: похоже, звук записался тихим или его не было.\n\n"
            "Проверьте, что во время урока звук шёл в колонки или наушники, "
            "а микрофон был включён и не был выключен на паузу (Mute)."
        )
    return text


# ============================================================
#  5. ОКНО ПРОГРАММЫ
# ============================================================


class LessonApp:
    """Простое окно: две кнопки, строка состояния и полоска прогресса."""

    def __init__(self, root):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry("640x260")
        self.root.minsize(560, 240)

        # Обмен сообщениями между рабочими потоками и окном.
        # Потоки только кладут сообщения в очередь, а окно читает их в своём потоке —
        # поэтому интерфейс никогда не подвисает.
        self.messages: queue.Queue = queue.Queue()

        self.state = "ready"                 # ready | recording | processing
        self.session: RecordingSession | None = None
        self.temp_dir: str | None = None
        self.mixed_wav: str | None = None
        self.recording_date: datetime | None = None
        self.recording_started_at = 0.0
        self.close_when_done = False
        self.reported_errors: set[str] = set()

        self.status_text = tk.StringVar(value=STATUS_READY)

        container = ttk.Frame(root, padding=20)
        container.pack(fill="both", expand=True)

        buttons = ttk.Frame(container)
        buttons.pack(pady=(6, 20))
        self.start_button = ttk.Button(buttons, text="Начать запись", width=22, command=self.on_start)
        self.start_button.pack(side="left", padx=8)
        self.stop_button = ttk.Button(buttons, text="Остановить", width=22,
                                      command=self.on_stop, state="disabled")
        self.stop_button.pack(side="left", padx=8)

        self.status_label = ttk.Label(container, textvariable=self.status_text,
                                      wraplength=580, justify="left", anchor="w")
        self.status_label.pack(fill="x")

        self.progress = ttk.Progressbar(container, mode="determinate", maximum=1.0)
        self.progress.pack(fill="x", pady=(16, 0))

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self._poll_messages)

    # ---------- маленькие помощники ----------

    def set_status(self, text: str) -> None:
        """Меняет текст в строке состояния (вызывается только из потока окна)."""
        self.status_text.set(text)
        log(text)

    def set_progress(self, value: float | None) -> None:
        """Показывает прогресс: число от 0 до 1 или None — «идёт работа, точного времени нет»."""
        if value is None:
            self.progress.stop()
            self.progress.configure(mode="indeterminate")
            self.progress.start(15)
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate", value=max(0.0, min(1.0, value)))

    def post(self, kind: str, payload=None) -> None:
        """Отправляет сообщение в окно из рабочего потока."""
        self.messages.put((kind, payload))

    # ---------- кнопки ----------

    def on_start(self) -> None:
        """Нажатие «Начать запись»: сначала проверяем устройства, затем пишем звук."""
        if self.state != "ready":
            return
        self.state = "checking"
        self.start_button.configure(state="disabled")
        self.set_progress(None)
        self.set_status("Проверка устройств...")
        # Проверка обращается к звуковой подсистеме и может занять пару секунд,
        # поэтому выполняем её в отдельном потоке — окно остаётся живым.
        threading.Thread(target=self._check_devices_worker, daemon=True).start()

    def _check_devices_worker(self) -> None:
        """Проверяет системный звук и микрофон в фоне."""
        com_initialized = init_com_for_thread()
        try:
            system_ok, system_info = check_system_audio()
            mic_ok, mic_info = check_microphone()
        finally:
            uninit_com_for_thread(com_initialized)
        self.post("devices_checked", (system_ok, system_info, mic_ok, mic_info))

    def _on_devices_checked(self, payload) -> None:
        """Показывает понятные предупреждения и решает, можно ли начинать запись."""
        system_ok, system_info, mic_ok, mic_info = payload

        if not system_ok and not mic_ok:
            self.state = "ready"
            self.start_button.configure(state="normal")
            self.set_progress(0.0)
            self.set_status("Ошибка: устройства звука недоступны")
            messagebox.showerror("Запись невозможна", f"{system_info}\n\n{mic_info}")
            return

        if not system_ok:
            if not messagebox.askyesno(
                "Системный звук недоступен",
                system_info + "\n\nПродолжить запись только с микрофоном (голос репетитора)?",
            ):
                self.state = "ready"
                self.start_button.configure(state="normal")
                self.set_progress(0.0)
                self.set_status(STATUS_READY)
                return

        if not mic_ok:
            if not messagebox.askyesno(
                "Микрофон недоступен",
                mic_info + "\n\nПродолжить запись только системного звука (голоса ученика)?",
            ):
                self.state = "ready"
                self.start_button.configure(state="normal")
                self.set_progress(0.0)
                self.set_status(STATUS_READY)
                return

        self.state = "ready"
        self.start_recording(use_system=system_ok, use_mic=mic_ok)

    def start_recording(self, use_system: bool = True, use_mic: bool = True) -> None:
        """Создаёт временную папку, запускает потоки записи и включает таймер."""
        cleanup_old_temp_dirs()
        self.temp_dir = tempfile.mkdtemp(prefix=TEMP_PREFIX)
        self.mixed_wav = None
        self.recording_date = datetime.now()
        self.recording_started_at = time.monotonic()
        self.reported_errors = set()
        self.session = RecordingSession(self.temp_dir, use_system, use_mic)

        self.state = "recording"
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.set_progress(None)
        self.set_status(f"{STATUS_RECORDING} 00:00:00")

        # Открытие устройств может занять пару секунд — делаем это в отдельном потоке,
        # чтобы окно не «замерло».
        threading.Thread(target=self._start_worker, args=(self.session,), daemon=True).start()

    def _start_worker(self, session: RecordingSession) -> None:
        session.start()
        problems = session.failed_sources()
        if len(problems) == len(session.sources) and problems:
            self.post("recording_error", "Не удалось начать запись.\n\n" + "\n".join(problems))
        elif problems:
            self.post("source_warning", "Одна из дорожек не записывается:\n\n" + "\n".join(problems))
        else:
            log("Запись началась.")

    def on_stop(self) -> None:
        """Нажатие «Остановить»: останавливаем запись и запускаем обработку."""
        if self.state != "recording":
            return
        self.state = "processing"
        self.stop_button.configure(state="disabled")
        self.set_status("Остановка записи...")
        self.set_progress(None)
        threading.Thread(target=self._process_worker, args=(self.session,), daemon=True).start()

    def _process_worker(self, session: RecordingSession) -> None:
        """Фоновая обработка: смешивание дорожек + распознавание речи."""
        try:
            session.stop()
            session.join()

            files = session.saved_files()
            if not files:
                self.post(
                    "error",
                    "Запись получилась пустой.\n\n"
                    "Если вы нажали «Остановить» сразу после «Начать запись» — "
                    "подождите пару секунд в следующий раз.\n\n"
                    "Если запись шла долго, проверьте, что звук урока слышен в колонках "
                    "или наушниках, а микрофон включён и не стоит на паузе (Mute).",
                )
                return

            duration = session.duration()
            log(f"Записано {format_hms(duration)}. Готовим звук к распознаванию...")
            self.post("status", "Обработка записи...")
            self.post("progress", None)

            # Если какое-то устройство записалось не в 16 кГц — приводим файл к 16 кГц.
            prepared = []
            for path in files:
                info = sf.info(path)
                if info.samplerate != SAMPLE_RATE:
                    converted = os.path.join(session.temp_dir, os.path.basename(path) + ".16k.wav")
                    log(f"Пересчитываем {os.path.basename(path)} "
                        f"({info.samplerate} Гц → {SAMPLE_RATE} Гц)...")
                    resample_wav(path, converted, SAMPLE_RATE)
                    prepared.append(converted)
                else:
                    prepared.append(path)

            mixed_path = os.path.join(session.temp_dir, "mixed.wav")
            peak, mixed_duration = mix_and_normalize(prepared, mixed_path)
            self.mixed_wav = mixed_path
            log(f"Смешано {format_hms(mixed_duration)}, пик {peak:.3f}.")

            if mixed_duration <= 0.5:
                self.post("error", "Запись получилась пустой (меньше секунды).")
                return

            if peak < SILENCE_PEAK:
                self.post("silent_recording", None)
                return

            text = transcribe_file(
                mixed_path,
                status_callback=lambda message: self.post("status", message),
                progress_callback=lambda value: self.post("progress", value),
            )

            folder = get_desktop_dir()
            file_name = f"Занятие {self.recording_date:%Y-%m-%d}"
            target = unique_path(folder, file_name, ".txt")
            with open(target, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(text + "\n")
            log(f"Транскрипция сохранена: {target}")

            self._remove_temp_dir()
            self.post("finished", target)
        except TranscriptionError as exc:
            self.post("transcribe_error", str(exc))
        except Exception as exc:  # noqa: BLE001 — показываем текст и пишем подробности в консоль
            log("Непредвиденная ошибка:\n" + traceback.format_exc())
            self.post("error", f"Неожиданная ошибка:\n{short_error(exc)}")

    def _remove_temp_dir(self) -> None:
        """Удаляет временные WAV-файлы (они больше не нужны)."""
        if self.temp_dir and os.path.isdir(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None
        self.mixed_wav = None

    # ---------- закрытие окна ----------

    def on_close(self) -> None:
        """Закрытие окна: предупреждаем, если идёт запись или обработка."""
        if self.state == "recording":
            if not messagebox.askyesno(
                "Идёт запись",
                "Запись ещё не остановлена. Выйти без транскрипции?\n\n"
                "Чтобы получить текст урока, нажмите «Нет», затем «Остановить».",
            ):
                return
            if self.session is not None:
                self.session.stop()
                self.session.join(timeout=5)
            self._remove_temp_dir()
            self.root.destroy()
            return

        if self.state == "processing":
            if not messagebox.askyesno(
                "Идёт обработка",
                "Сейчас идёт обработка записи. Если закрыть программу сейчас, "
                "транскрипция не сохранится.\n\nЗакрыть всё равно?",
            ):
                return
            if messagebox.askyesno(
                "Подождать обработку?",
                "Закрыть программу автоматически, когда транскрипция будет сохранена?",
            ):
                self.close_when_done = True
                self.set_status(self.status_text.get() + "  (окно закроется само)")
                return

        self.root.destroy()

    # ---------- общение с рабочими потоками ----------

    def _poll_messages(self) -> None:
        """Каждые 0,1 секунды проверяет сообщения от потоков и обновляет таймер."""
        try:
            while True:
                kind, payload = self.messages.get_nowait()
                try:
                    self._handle_message(kind, payload)
                except Exception:  # noqa: BLE001 — окно не должно падать из-за одной ошибки
                    log("Ошибка при обновлении окна:\n" + traceback.format_exc())
        except queue.Empty:
            pass

        if self.state == "recording":
            elapsed = time.monotonic() - self.recording_started_at
            self.status_text.set(f"{STATUS_RECORDING} {format_hms(elapsed)}")
            self._check_source_errors()

        if self.close_when_done and self.state == "ready":
            self.close_when_done = False
            self.root.destroy()
            return

        self.root.after(100, self._poll_messages)

    def _check_source_errors(self) -> None:
        """Сообщает, если одна из дорожек перестала писаться во время урока."""
        if self.session is None:
            return
        for source in self.session.sources:
            if source.error is None or source.display_name in self.reported_errors:
                continue
            self.reported_errors.add(source.display_name)
            messagebox.showwarning(
                "Проблема с записью",
                f"{source.display_name}: запись прервалась.\n{short_error(source.error)}\n\n"
                "Вторая дорожка продолжает записываться. Вы можете остановить запись "
                "и сохранить то, что уже есть.",
            )

    def _handle_message(self, kind: str, payload) -> None:
        """Реакция окна на сообщение из рабочего потока."""
        if kind == "devices_checked":
            self._on_devices_checked(payload)
        elif kind == "status":
            self.set_status(str(payload))
        elif kind == "progress":
            self.set_progress(None if payload is None else float(payload))
        elif kind == "finished":
            self.state = "ready"
            self.start_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            self.set_progress(0.0)
            self.set_status(f"Транскрипция готова: {payload}")
            messagebox.showinfo("Готово!", f"Транскрипция сохранена:\n\n{payload}")
        elif kind == "silent_recording":
            self._reset_to_ready()
            self.set_status("Ошибка: запись пустая — звука не было")
            messagebox.showerror(
                "Запись пустая",
                "В записи не оказалось звука, поэтому транскрипция не создана.\n\n"
                + ADVICE_SYSTEM_SOUND,
            )
        elif kind == "recording_error":
            self._reset_to_ready()
            self.set_status("Ошибка: не удалось начать запись")
            self._remove_temp_dir()
            messagebox.showerror("Не удалось начать запись", str(payload))
        elif kind == "source_warning":
            messagebox.showwarning("Проблема с записью", str(payload))
        elif kind == "transcribe_error":
            self._offer_raw_audio(str(payload))
        elif kind == "error":
            self._reset_to_ready()
            self.set_status(f"Ошибка: {str(payload).splitlines()[0]}")
            messagebox.showerror("Ошибка", str(payload))
        elif kind == "audio_saved":
            self.set_status(f"Аудио сохранено: {payload}")
        elif kind == "audio_saved_dialog":
            self._show_audio_saved_dialog(payload)
        else:  # неизвестное сообщение — просто записываем в журнал
            log(f"Неизвестное сообщение: {kind}")

    def _show_audio_saved_dialog(self, payload) -> None:
        """Сообщает пользователю, что запись сохранена на рабочий стол в формате WAV."""
        target, message = payload
        messagebox.showinfo(
            "Аудио сохранено",
            "Запись сохранена на рабочий стол:\n\n"
            f"{target}\n\n"
            "Файл можно открыть любым плеером. Как распознать его позже — написано "
            "в конце файла README.md.\n\n"
            "Ошибка распознавания была такой:\n" + str(message)[:500],
        )

    def _reset_to_ready(self) -> None:
        """Возвращает окно в исходное состояние (кнопки как в начале)."""
        self.state = "ready"
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.set_progress(0.0)
        self.session = None

    def _offer_raw_audio(self, message: str) -> None:
        """
        Если распознать текст не удалось — предлагаем сохранить запись (WAV),
        чтобы можно было попробовать позже.
        """
        self._reset_to_ready()
        first_line = message.splitlines()[0] if message else "Ошибка распознавания"
        self.set_status(f"Ошибка: {first_line}")

        if not self.mixed_wav or not os.path.exists(self.mixed_wav):
            messagebox.showerror("Ошибка распознавания", message)
            return

        if not messagebox.askyesno(
            "Не удалось распознать текст",
            message + "\n\nСохранить запись на рабочий стол, чтобы попробовать ещё раз позже?",
        ):
            folder = self.temp_dir or ""
            self.set_status(f"Ошибка: {first_line}")
            if folder:
                messagebox.showinfo(
                    "Запись сохранена временно",
                    "Аудио осталось здесь:\n\n" f"{folder}\n\n"
                    "Папка удалится, когда вы закроете программу.",
                )
            return

        threading.Thread(target=self._save_raw_audio_worker, args=(message,), daemon=True).start()

    def _save_raw_audio_worker(self, message: str) -> None:
        """Копирует запись на рабочий стол в фоне (файл может быть большим)."""
        try:
            self.post("status", "Сохранение аудио...")
            folder = get_desktop_dir()
            base = f"Занятие {self.recording_date:%Y-%m-%d} (аудио)"
            target = unique_path(folder, base, ".wav")
            shutil.copyfile(self.mixed_wav, target)
            log(f"Аудио сохранено: {target}")
            self.post("audio_saved", target)
            self.post("audio_saved_dialog", (target, message))
        except Exception as exc:  # noqa: BLE001
            self.post("error", f"Не удалось сохранить аудио:\n{short_error(exc)}")


def transcribe_existing_file(wav_path: str) -> int:
    """
    Распознаёт уже готовый WAV-файл из командной строки.

    Пригодится, если первая попытка распознавания не удалась, а запись
    сохранилась на рабочем столе:

        python main.py "C:\\Users\\Имя\\Desktop\\Занятие 2026-10-13 (аудио).wav"
    """
    if not os.path.isfile(wav_path):
        print(f"Файл не найден: {wav_path}")
        return 1

    base_name = os.path.splitext(os.path.basename(wav_path))[0]

    try:
        info = sf.info(wav_path)
    except Exception as exc:  # noqa: BLE001
        print(
            "Не удалось прочитать этот звуковой файл.\n"
            f"Подробности: {short_error(exc)}\n\n"
            "Поддерживаются файлы WAV, MP3, FLAC и OGG."
        )
        return 1

    if info.samplerate != SAMPLE_RATE:
        print(f"Приводим файл к {SAMPLE_RATE} Гц (это может занять минуту)...")
        converted = os.path.join(
            tempfile.mkdtemp(prefix=TEMP_PREFIX), base_name + ".16k.wav"
        )
        resample_wav(wav_path, converted, SAMPLE_RATE)
        wav_path = converted

    printed = {"value": -1.0}

    def show_progress(value: float) -> None:
        percent = int(value * 100)
        if percent >= printed["value"] + 5:
            printed["value"] = percent
            print(f"\rРаспознано: {percent}%", end="", flush=True)

    try:
        text = transcribe_file(wav_path, status_callback=print, progress_callback=show_progress)
    except TranscriptionError as exc:
        print(f"\nОшибка: {exc}")
        return 1

    folder = get_desktop_dir()
    target = unique_path(folder, base_name, ".txt")
    with open(target, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(text + "\n")
    print(f"\nГотово! Текст сохранён: {target}")
    return 0


def main() -> int:
    """Точка входа: создаём окно и запускаем его цикл обработки событий."""
    # Если в командной строке передали файл — просто распознаём его без окна.
    for argument in sys.argv[1:]:
        if argument.lower().endswith((".wav", ".flac", ".ogg", ".mp3")) or os.path.isfile(argument):
            return transcribe_existing_file(argument)

    if tk is None:
        print(
            "В вашей установке Python нет модуля tkinter, без него окно программы не построить.\n"
            "Переустановите Python с сайта python.org (модуль tkinter входит в установщик\n"
            "по умолчанию) и запустите программу снова."
        )
        return 1

    # Отключаем предупреждение Hugging Face про символические ссылки в Windows —
    # новичку оно ни о чём не говорит, а модель всё равно скачается нормально.
    os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")

    root = tk.Tk()
    try:  # аккуратный внешний вид в Windows 10/11
        ttk.Style().theme_use("vista")
    except Exception:
        pass
    LessonApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nПрограмма закрыта.")
