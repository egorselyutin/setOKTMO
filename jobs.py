"""Фоновые задания обработки. Модуль не зависит от Streamlit.

ЗАЧЕМ ЭТО НУЖНО
---------------
Streamlit выполняет скрипт страницы заново при любом «перезапуске»: клик по странице, обрыв и
восстановление соединения, перегрузка сервера. Если обработка файла идёт ВНУТРИ скрипта, такой
перезапуск обрывает её на середине (а в старой версии ещё и приводил к ошибке KeyError: 'result').
На своём компьютере этого почти не бывает, а на сервере с долгой обработкой — постоянно.

Теперь обработка идёт в отдельном потоке на сервере и живёт независимо от скрипта страницы.
Страница лишь показывает ход выполнения; после «перезапуска» она просто подключается к тому же
заданию и продолжает показывать прогресс — работа не теряется и не запускается заново.

Заодно задания выстраиваются в очередь (по умолчанию одно за раз): на небольшом сервере несколько
одновременных обработок крупных файлов исчерпали бы память.
"""

from __future__ import annotations

import os
import threading
import time
import traceback


class Cancelled(BaseException):
    """Задание отменено (пользователь убрал файл). BaseException — чтобы её не перехватывали
    обычные `except Exception` внутри конвейера."""


class Job:
    """Состояние одного задания. Поля обновляются из потока задания, читаются страницей."""

    def __init__(self, key: str, session: str):
        self.key = key
        self.session = session
        self.created = time.time()
        self.started: float | None = None   # когда обработка реально началась (после ожидания очереди)
        self.finished: float | None = None
        self.queued = True                  # True, пока задание ждёт своей очереди
        self.progress = (0.0, "Подготовка…", 0, 0)  # (доля 0..1, название этапа, сделано, всего)
        self.result = None
        self.error: str | None = None
        self.done = False
        self.cancelled = False

    def report(self, frac, stage, done, total):
        """Колбэк прогресса для конвейера. Заодно — точка, где можно остановить задание."""
        if self.cancelled:
            raise Cancelled()
        self.progress = (frac, stage, done, total)


class JobManager:
    """Реестр заданий. Один экземпляр на весь сервер (создаётся через st.cache_resource)."""

    def __init__(self, max_concurrent: int | None = None):
        if max_concurrent is None:
            env = os.environ.get("OKTMO_MAX_JOBS", "")
            max_concurrent = int(env) if env.isdigit() and int(env) > 0 else 1
        self.max_concurrent = max_concurrent
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()
        self._slots = threading.BoundedSemaphore(max_concurrent)

    # --- чтение -----------------------------------------------------------------------------
    def get(self, key: str) -> Job | None:
        with self._lock:
            return self._jobs.get(key)

    def waiting_ahead(self) -> int:
        """Сколько заданий сейчас ждут очереди (для сообщения пользователю)."""
        with self._lock:
            return sum(1 for j in self._jobs.values() if j.queued and not j.done)

    # --- запуск и отмена ------------------------------------------------------------------
    def start(self, key: str, session: str, target) -> Job:
        """Запускает target(job) в фоновом потоке и возвращает Job.

        Если задание с таким ключом уже есть (страница перезапустилась) — возвращает его, а не
        запускает заново. Остальные задания той же сессии отменяются (пользователь загрузил другой файл).
        """
        with self._lock:
            existing = self._jobs.get(key)
            if existing is not None:
                return existing
            self._cancel_locked(session, keep=key)
            job = Job(key, session)
            self._jobs[key] = job
        threading.Thread(target=self._run, args=(job, target), daemon=True, name="oktmo-job").start()
        return job

    def cancel_session(self, session: str, keep: str | None = None):
        """Отменяет и забывает все задания сессии (например, файл убрали крестиком)."""
        with self._lock:
            self._cancel_locked(session, keep)

    def _cancel_locked(self, session: str, keep: str | None):
        for key, job in list(self._jobs.items()):
            if job.session == session and key != keep:
                job.cancelled = True
                del self._jobs[key]  # поток сам завершится при ближайшем отчёте о прогрессе

    def purge(self, max_age: float = 1800.0):
        """Освобождает память: удаляет завершённые задания, которым больше max_age секунд."""
        now = time.time()
        with self._lock:
            for key, job in list(self._jobs.items()):
                if job.done and job.finished is not None and now - job.finished > max_age:
                    del self._jobs[key]

    # --- поток задания -------------------------------------------------------------------
    def _run(self, job: Job, target):
        try:
            with self._slots:  # ждём своей очереди
                if job.cancelled:
                    return
                job.queued = False
                job.started = time.time()
                job.progress = (0.0, "Подготовка…", 0, 0)
                job.result = target(job)
        except Cancelled:
            pass
        except BaseException as e:  # noqa: BLE001 — любая причина должна дойти до пользователя
            job.error = f"Не удалось обработать файл: {e}"
            traceback.print_exc()  # полный трейсбек — в логи сервера («Manage app» -> Logs)
        finally:
            job.finished = time.time()
            job.done = True
