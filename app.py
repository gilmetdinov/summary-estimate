"""
Десктопное окно конвертера смет ПИР (tkinter).

Сметчица: двойной клик по exe (или `python3 app.py`) -> окно -> выбрать выгрузку(и)
ГС, поправить реквизиты/ФИО, при необходимости включить строки изысканий/госэкспертизы
-> «Конвертировать» -> указать, куда сохранить -> готовый комплект (.xlsx, а для
нескольких выгрузок — отдельный файл на каждую в выбранной папке).

tkinter входит в стандартную поставку Python с python.org (Windows/macOS), отдельно
ставить не нужно. Вся «тяжёлая» логика — в config.py / smeta_render.py; здесь только UI,
поэтому конфиг (config_from_strings) и конвертация покрыты тестами без запуска окна.
"""
from __future__ import annotations
import os
import subprocess
import sys
import threading

# В собранном «оконном» exe (--noconsole) sys.stdout/stderr = None, и любой print()
# или предупреждение упадёт с AttributeError. Подменяем заглушкой ещё до импортов.
for _name in ("stdout", "stderr"):
    if getattr(sys, _name, None) is None:
        setattr(sys, _name, open(os.devnull, "w"))

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:  # pragma: no cover
    raise SystemExit(
        "Не найден tkinter. Используйте Python с python.org "
        "(в его стандартной поставке tkinter есть).")

from config import load_config, config_from_strings, Config
from smeta_render import convert, convert_batch, suggest_name

APP_TITLE = "Конвертер смет ПИР"


def resource_path(name: str) -> str:
    """Путь к ресурсу: в собранном exe — из _MEIPASS, иначе рядом со скриптом."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def default_config() -> Config:
    """Дефолты формы: config.json рядом с exe, иначе вложенный, иначе встроенные."""
    for p in ("config.json", resource_path("config.json")):
        if os.path.exists(p):
            return load_config(p)
    return Config()


DONOR_PATH = resource_path("template.xlsx")


def open_folder(path: str) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=False)
        else:
            subprocess.run(["xdg-open", path], check=False)
    except Exception:
        pass


class App:
    # поля: (ключ, подпись). Группируются по рамкам.
    ORG_FIELDS = [
        ("org_short", "Наименование организации"),
        ("org_dir_title1", "Должность, строка 1"),
        ("org_dir_title2", "Должность, строка 2"),
        ("org_dir_name", "ФИО руководителя"),
        ("gip_name", "ФИО ГИП"),
    ]
    CUST_FIELDS = [
        ("cust_dir_title", "Должность, строка 1"),
        ("cust_dir_title2", "Должность, строка 2 / филиал"),
        ("cust_dir_name", "ФИО руководителя"),
    ]

    def __init__(self, root: "tk.Tk"):
        self.root = root
        self.cfg = default_config()
        self.input_paths: list[str] = []
        self.entries: dict[str, "tk.StringVar"] = {}
        self.izysk_on = tk.BooleanVar(value=False)
        self.gosexp_on = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="")
        self.files_label = tk.StringVar(value="Файлы не выбраны")

        root.title(APP_TITLE)
        set_window_icon(root)
        root.minsize(560, 560)
        self._build()
        # подгоняем окно ровно под содержимое (корректно при любом DPI-масштабе)
        root.update_idletasks()
        root.geometry(f"{root.winfo_reqwidth()}x{root.winfo_reqheight()}")

    # ---------- построение интерфейса ----------
    def _build(self):
        pad = {"padx": 10, "pady": 4}
        ttk.Label(self.root, text=APP_TITLE, font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(10, 0))
        ttk.Label(self.root, wraplength=600, foreground="#555",
                  text="Выгрузка из Гранд Сметы / Адепта (.xlsx) → фирменный комплект "
                       "(листы смет + сводная). Можно несколько выгрузок сразу.").pack(anchor="w", padx=12)

        # 1. файлы
        f1 = ttk.LabelFrame(self.root, text="1. Выгрузка(и) ГС")
        f1.pack(fill="x", **pad)
        ttk.Button(f1, text="Выбрать файлы…", command=self._pick_files).pack(side="left", padx=8, pady=8)
        ttk.Label(f1, textvariable=self.files_label, foreground="#2563eb",
                  wraplength=420).pack(side="left", padx=4)

        # 2. проектная организация
        f2 = ttk.LabelFrame(self.root, text="2. Проектная организация (подписант)")
        f2.pack(fill="x", **pad)
        self._fields(f2, self.ORG_FIELDS)

        # 3. заказчик
        f3 = ttk.LabelFrame(self.root, text="3. Заказчик (подписант)")
        f3.pack(fill="x", **pad)
        self._fields(f3, self.CUST_FIELDS)

        # 4. параметры + ручные строки
        f4 = ttk.LabelFrame(self.root, text="4. Параметры и ручные строки сводной")
        f4.pack(fill="x", **pad)
        row = ttk.Frame(f4); row.pack(fill="x", padx=8, pady=4)
        ttk.Label(row, text="Ставка НДС, %").pack(side="left")
        v = tk.StringVar(value=self.cfg.vat_pct); self.entries["vat_pct"] = v
        ttk.Entry(row, textvariable=v, width=8).pack(side="left", padx=8)

        self._placeholder(f4, self.izysk_on, "izysk",
                          "Добавить строку изысканий (колонка «инж. изысканий»)",
                          name_default="Инженерно-геодезические изыскания",
                          osn_default="НЗ_ИГДИ")
        self._placeholder(f4, self.gosexp_on, "gosexp",
                          "Добавить строку госэкспертизы (колонка «проектных работ»)",
                          name_default="Государственная экспертиза проектной документации "
                                       "в части проверки достоверности определения сметной стоимости",
                          osn_default="")

        # поля ед. стоимости (колонка O сводной, для расчёта стоимости за единицу в колонке P)
        unit_frame = ttk.Frame(f4); unit_frame.pack(fill="x", padx=8, pady=4)
        ttk.Label(unit_frame, text="Ед. стоимость:", font=("", 9, "bold")).pack(anchor="w")
        for key, caption, default in [
            ("unit_cost_proj", "проектных работ", ""),
            ("unit_cost_izysk", "изысканий", ""),
            ("unit_cost_gosexp", "госэкспертизы", ""),
        ]:
            row = ttk.Frame(unit_frame); row.pack(fill="x", pady=2)
            ttk.Label(row, text=caption, width=26, anchor="w").pack(side="left")
            var = tk.StringVar(value=default); self.entries[key] = var
            ttk.Entry(row, textvariable=var, width=12).pack(side="left")
            ttk.Label(row, text="тыс.₽ / ед.", foreground="#888").pack(side="left", padx=4)

        # кнопка + статус
        bar = ttk.Frame(self.root); bar.pack(fill="x", padx=12, pady=10)
        self.btn = ttk.Button(bar, text="Конвертировать", command=self.on_convert)
        self.btn.pack(side="left")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(side="left", padx=10)

    def _fields(self, parent, fields):
        for key, caption in fields:
            row = ttk.Frame(parent); row.pack(fill="x", padx=8, pady=2)
            ttk.Label(row, text=caption, width=26, anchor="w").pack(side="left")
            var = tk.StringVar(value=getattr(self.cfg, key))
            self.entries[key] = var
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

    def _placeholder(self, parent, flag_var, prefix, caption, name_default, osn_default):
        ph = getattr(self.cfg, "izyskaniya" if prefix == "izysk" else "gosexpertiza")
        name0 = ph.name if ph else name_default
        osn0 = ph.osn if ph else osn_default
        val0 = "" if not ph or ph.value is None else f"{ph.value:g}"
        if ph and ph.value is not None:
            flag_var.set(True)
        ttk.Checkbutton(parent, text=caption, variable=flag_var).pack(anchor="w", padx=8, pady=(8, 0))
        row = ttk.Frame(parent); row.pack(fill="x", padx=24, pady=2)
        for key, cap, default, width in [
            (f"{prefix}_name", "Наименование", name0, None),
            (f"{prefix}_osn", "Обоснование", osn0, 16),
            (f"{prefix}_value", "Стоимость, тыс.₽", val0, 12),
        ]:
            ttk.Label(row, text=cap).pack(side="left", padx=(0, 4))
            var = tk.StringVar(value=default); self.entries[key] = var
            e = ttk.Entry(row, textvariable=var, width=width) if width else ttk.Entry(row, textvariable=var)
            e.pack(side="left", fill="x", expand=(width is None), padx=(0, 8))

    # ---------- действия ----------
    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите выгрузку(и) ГС", filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")])
        if paths:
            self.input_paths = list(paths)
            names = ", ".join(os.path.basename(p) for p in self.input_paths)
            self.files_label.set(f"Выбрано {len(self.input_paths)}: {names}")

    def _collect(self) -> dict:
        values = {k: v.get() for k, v in self.entries.items()}
        values["izysk_on"] = self.izysk_on.get()
        values["gosexp_on"] = self.gosexp_on.get()
        return values

    def on_convert(self):
        if not self.input_paths:
            messagebox.showwarning(APP_TITLE, "Сначала выберите выгрузку(и) ГС.")
            return
        cfg = config_from_strings(self._collect(), default_config())
        if len(self.input_paths) == 1:
            out = filedialog.asksaveasfilename(
                title="Сохранить результат", defaultextension=".xlsx",
                initialfile=suggest_name(self.input_paths[0]), filetypes=[("Excel", "*.xlsx")])
            if not out:
                return
            target = ("single", out)
        else:
            out_dir = filedialog.askdirectory(title="Папка для результатов")
            if not out_dir:
                return
            target = ("batch", out_dir)
        self._run_async(cfg, target)

    def _run_async(self, cfg, target):
        self.btn.config(state="disabled")
        self.status.set("Конвертирую…")

        def work():
            try:
                kind, dest = target
                if kind == "single":
                    convert(self.input_paths[0], dest, DONOR_PATH, cfg)
                    msg, folder = f"Готово:\n{dest}", os.path.dirname(dest)
                else:
                    res = convert_batch(self.input_paths, dest, DONOR_PATH, cfg)
                    ok = sum(1 for r in res if r["error"] is None)
                    msg = f"Готово: {ok} из {len(res)} в\n{dest}"
                    errs = [f"• {os.path.basename(r['input'])}: {r['error']}"
                            for r in res if r["error"]]
                    if errs:
                        msg += "\n\nНе удалось:\n" + "\n".join(errs)
                    folder = dest
                self.root.after(0, lambda: self._done(True, msg, folder))
            except Exception as ex:  # noqa: BLE001 — показать сметчице понятную ошибку
                self.root.after(0, lambda e=ex: self._done(False, str(e), None))

        threading.Thread(target=work, daemon=True).start()

    def _done(self, ok, msg, folder):
        self.btn.config(state="normal")
        self.status.set("")
        if ok:
            if messagebox.askyesno(APP_TITLE, msg + "\n\nОткрыть папку с результатом?"):
                open_folder(folder)
        else:
            messagebox.showerror("Ошибка конвертации", msg)


def _windows_init():
    """Windows-инициализация ДО создания окна:
    - DPI-awareness (иначе интерфейс мыльный — система растягивает картинку битмапом);
    - AppUserModelID — иначе в панели задач висит стандартная иконка Python, а не наша.
    На других ОС безопасно игнорируется."""
    if not sys.platform.startswith("win"):
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # система (Win 8.1+)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()     # запасной (Win Vista+)
        except Exception:
            pass
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "neuroestimate.smeta.converter")
    except Exception:
        pass


def set_window_icon(root):
    """Иконка окна и панели задач. PyInstaller icon= красит только сам файл .exe —
    окну иконку надо задавать в коде."""
    ico = resource_path("app.ico")
    try:
        if os.path.exists(ico):
            root.iconbitmap(default=ico)   # Windows: окно + таскбар
            return
    except Exception:
        pass
    png = resource_path("app_icon.png")    # фолбэк (macOS/Linux)
    try:
        if os.path.exists(png):
            root._icon_img = tk.PhotoImage(file=png)  # держим ссылку, иначе GC съест
            root.iconphoto(True, root._icon_img)
    except Exception:
        pass


def main():
    _windows_init()
    root = tk.Tk()
    # масштаб шрифтов под реальный DPI -> текст крупный и чёткий на HiDPI
    try:
        root.tk.call("tk", "scaling", root.winfo_fpixels("1i") / 72.0)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:  # без консоли трейсбек не виден — покажем окном и в лог
        try:
            with open("summary-estimate-error.log", "w", encoding="utf-8") as f:
                import traceback
                traceback.print_exc(file=f)
        except Exception:
            pass
        try:
            from tkinter import messagebox
            messagebox.showerror("Ошибка запуска", str(ex))
        except Exception:
            pass
        raise
