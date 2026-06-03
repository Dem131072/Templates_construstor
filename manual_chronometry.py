"""
Инструмент проведения хронометража заполнения шаблонов документов DOCX вручную через Microsoft Word.
Подсчитывает время, затраченное на заполнение шаблона.
Производит сравнение исходного и выходного документа.
Записывает полученные данные в файл логов.

Пользователь выбирает шаблон и указывает имя нового файла.
Открывается новый файл, в котором производятся замены.
После завершения заполнения пользователь сохраняет новый файл и завершает работу программы.
"""


import os
import re
import shutil
import subprocess
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import filedialog, messagebox

WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
P, T, TAB, BR, CR = (WORD_NS + x for x in ("p", "t", "tab", "br", "cr"))
DOCX_PART = "word/document.xml"
WORD_RE = re.compile(r"[A-Za-zА-Яа-яЁё0-9]+(?:[-'][A-Za-zА-Яа-яЁё0-9]+)?")
WINDOW_TITLE = "DOCX хронометраж"
WINDOW_SIZE = "720x245"
LOG_FILE = "logs.txt"

@dataclass(frozen=True)
class Report:
    template: str
    symbols: int
    words: int
    tokens: int
    replacements: int
    timing: float


def program_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
    

def paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for element in paragraph.iter():
        if element.tag == T and element.text:
            parts.append(element.text)
        elif element.tag in (TAB, BR, CR):
            parts.append(" ")
    return "".join(parts)


def read_docx_text(path: Path) -> str:
    chunks = []
    with zipfile.ZipFile(path) as docx:
        root = ET.fromstring(docx.read(DOCX_PART))
        for paragraph in root.iter(P):
            text = paragraph_text(paragraph)
            if text.strip():
                chunks.append(text)
    return " ".join(chunks)


def count_replacements(before_text: str, after_text: str) -> int:
    before, after = before_text.split(), after_text.split()
    diff = SequenceMatcher(None, before, after, autojunk=False)
    return sum(1 for tag, _i1, _i2, _j1, _j2 in diff.get_opcodes() if tag != "equal")


def analyze(source: Path, output: Path, elapsed: float) -> Report:
    source_text = read_docx_text(source)
    output_text = read_docx_text(output)
    return Report(
        template=source.name,
        symbols=len(output_text),
        words=len(WORD_RE.findall(output_text)),
        tokens=len(output_text.split()),
        replacements=count_replacements(source_text, output_text),
        timing=round(elapsed, 3),
    )


def wait_file_stable(path: Path, timeout: float = 15.0) -> None:
    deadline = time.time() + timeout
    last_signature = None
    stable_checks = 0
    while time.time() < deadline:
        try:
            stat = path.stat()
            signature = (stat.st_size, stat.st_mtime_ns)
            with path.open("rb"):
                pass
        except OSError:
            stable_checks = 0
        else:
            stable_checks = stable_checks + 1 if signature == last_signature else 0
            last_signature = signature
            if stable_checks >= 3:
                return
        time.sleep(0.35)


def write_log(report: Report) -> None:
    lines = [
        "------------------------------------------",
        "Заполнение документа через Microsoft Word",
        f"Дата={datetime.now().isoformat(timespec='seconds')}",
        f"Документ={report.template}",
        f"Всего символов={report.symbols}",
        f"Всего слов={report.words}",
        f"Всего токенов={report.tokens}",
        f"Замены={report.replacements}",
        f"Время заполнения={report.timing}",
        "------------------------------------------"
    ]
    with (program_dir() / LOG_FILE).open("a", encoding="utf-8") as file:
        file.write("\n".join(lines))
        

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.source_path = None
        self.output_path = None
        self.started_at = None
        root.title(WINDOW_TITLE)
        root.geometry(WINDOW_SIZE)
        root.resizable(False, False)
        self.show_start_screen()

    def clear(self) -> None:
        for widget in self.root.winfo_children():
            widget.destroy()

    def path_row(self, row: int, label: str, command) -> tk.Entry:
        tk.Label(self.root, text=label).grid(row=row, column=0, padx=16, pady=14, sticky="w")
        entry = tk.Entry(self.root, width=70)
        entry.grid(row=row, column=1, padx=8, pady=14)
        tk.Button(self.root, text="Выбрать файл", command=command).grid(row=row, column=2, padx=8, pady=14)
        return entry

    def set_entry(self, entry: tk.Entry, path: Path) -> None:
        entry.delete(0, tk.END)
        entry.insert(0, str(path))

    def show_start_screen(self) -> None:
        self.clear()
        self.source_entry = self.path_row(0, "Исходный файл .docx", self.choose_source)
        self.output_entry = self.path_row(1, "Итоговый файл .docx", self.choose_output)
        tk.Label(self.root, text="Хронометраж заполнения документов Word",  font=("Times New Roman", 15)).grid(row=2, column=0, columnspan=3, pady=12)
        tk.Button(self.root, text="Начать заполнение", width=24, command=self.start).grid(row=3, column=0, columnspan=3, pady=12)

    def choose_source(self) -> None:
        filename = filedialog.askopenfilename(title="Выберите исходный DOCX", filetypes=[("Word документы", "*.docx")])
        if not filename:
            return
        self.source_path = Path(filename)
        self.output_path = self.source_path.with_name(self.source_path.stem + "_filled.docx")
        self.set_entry(self.source_entry, self.source_path)
        self.set_entry(self.output_entry, self.output_path)

    def choose_output(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Укажите итоговый DOCX",
            defaultextension=".docx",
            initialdir=str(self.source_path.parent) if self.source_path else str(Path.cwd()),
            initialfile=self.source_path.stem + "_filled.docx" if self.source_path else "output.docx",
            filetypes=[("Word документы", "*.docx")],
        )
        if filename:
            self.output_path = Path(filename)
            self.set_entry(self.output_entry, self.output_path)

    def validate(self, source: Path, output: Path) -> bool:
        checks = [
            (source.exists(), "Исходный файл не найден."),
            (source.suffix.lower() == ".docx", "Поддерживаются только .docx файлы."),
            (output.suffix.lower() == ".docx", "Итоговый файл должен быть .docx."),
            (not output.exists(), "Итоговый файл уже существует. Укажите другое имя."),
        ]
        for ok, message in checks:
            if not ok:
                messagebox.showerror("Ошибка", message)
                return False
        return True

    def start(self) -> None:
        """Запуск на Windows"""
        try:
            source = Path(self.source_entry.get().strip())
            output = Path(self.output_entry.get().strip())
            if output.suffix.lower() != ".docx":
                output = output.with_suffix(".docx")
            if not self.validate(source, output):
                return
            output.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, output)
            self.source_path = source
            self.output_path = output
            self.started_at = time.time()
            os.startfile(str(output))
            self.show_work_screen()
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))

    def show_work_screen(self) -> None:
        self.clear()
        tk.Label(self.root, text="1.Сохраните заполненный документ\n2.Нажмите 'Готово'", 
                 font=("Arial", 20), wraplength=700, justify="center").pack(pady=55)
        tk.Button(self.root, text="Готово", font=("Arial", 20), width=24, height=2, command=self.finish).pack()

    def finish(self) -> None:
        try:
            if self.started_at is None or self.source_path is None or self.output_path is None:
                messagebox.showerror("Ошибка", "Замер не был начат.")
                return
            elapsed = time.time() - self.started_at
            wait_file_stable(self.output_path)
            report = analyze(self.source_path, self.output_path, elapsed)
            write_log(report)
            messagebox.showinfo(
                "Готово",
                "Благодарю за помощь!\n"
                f"Документ: {report.template}\n"
                f"Символов: {report.symbols}\n"
                f"Слов: {report.words}\n"
                f"Токенов: {report.tokens}\n"
                f"Замен: {report.replacements}\n"
                f"Время: {report.timing} сек.",
            )
            self.root.destroy()
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))


def main() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
