import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import re
import json

# ──────────────────────────────────────────────
#  Зависимости с проверкой
# ──────────────────────────────────────────────
missing = []
try:
    from yandex_music import Client as YMClient
except ImportError:
    missing.append("yandex_music")
try:
    from ytmusicapi import YTMusic
except ImportError:
    missing.append("ytmusicapi")
try:
    import openpyxl
except ImportError:
    missing.append("openpyxl")

# ══════════════════════════════════════════════
#  Цветовая схема
# ══════════════════════════════════════════════
BG        = "#0f0f13"
BG2       = "#1a1a24"
BG3       = "#22222f"
ACCENT    = "#6c63ff"
ACCENT2   = "#ff6584"
FG        = "#e8e8f0"
FG2       = "#9090a8"
SUCCESS   = "#43e97b"
WARNING   = "#f9ca24"
ERROR     = "#ff6b6b"
FONT_MAIN = ("Segoe UI", 10)
FONT_H    = ("Segoe UI Semibold", 11)
FONT_BIG  = ("Segoe UI Bold", 13)

# ══════════════════════════════════════════════
#  Главное окно
# ══════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Music Transfer  ·  Яндекс → YouTube")
        self.geometry("860x640")
        self.minsize(760, 560)
        self.configure(bg=BG)
        self._style()
        self._build()
        if missing:
            self.after(300, self._warn_missing)

    # ── ttk-стили ─────────────────────────────
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TNotebook",        background=BG,  borderwidth=0)
        s.configure("TNotebook.Tab",    background=BG3, foreground=FG2,
                    font=FONT_H, padding=[18, 8])
        s.map("TNotebook.Tab",
              background=[("selected", BG2)],
              foreground=[("selected", FG)])
        s.configure("TFrame",           background=BG2)
        s.configure("TLabel",           background=BG2, foreground=FG,  font=FONT_MAIN)
        s.configure("TEntry",           fieldbackground=BG3, foreground=FG,
                    insertcolor=FG, borderwidth=0, font=FONT_MAIN)
        s.configure("Accent.TButton",   background=ACCENT, foreground="#fff",
                    font=FONT_H, borderwidth=0, padding=[12, 7])
        s.map("Accent.TButton",
              background=[("active", "#8078ff"), ("disabled", BG3)])
        s.configure("Red.TButton",      background=ACCENT2, foreground="#fff",
                    font=FONT_H, borderwidth=0, padding=[12, 7])
        s.map("Red.TButton",
              background=[("active", "#ff8fa3"), ("disabled", BG3)])
        s.configure("TProgressbar",     troughcolor=BG3, background=ACCENT,
                    borderwidth=0, thickness=8)
        s.configure("TSeparator",       background=BG3)

    # ── Шапка ─────────────────────────────────
    def _build(self):
        hdr = tk.Frame(self, bg=BG, pady=14)
        hdr.pack(fill="x", padx=24)
        tk.Label(hdr, text="🎵 Music Transfer", bg=BG, fg=FG,
                 font=("Segoe UI Bold", 16)).pack(side="left")
        tk.Label(hdr, text="Яндекс Музыка  →  YouTube Music",
                 bg=BG, fg=FG2, font=FONT_MAIN).pack(side="left", padx=14)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(0,12))

        self.tab_ym  = YandexTab(nb)
        self.tab_yt  = YouTubeTab(nb)
        nb.add(self.tab_ym,  text="  ❶  Яндекс Музыка  ")
        nb.add(self.tab_yt,  text="  ❷  YouTube Music  ")

    def _warn_missing(self):
        msg = "Установи зависимости:\n\npip install " + " ".join(missing)
        messagebox.showwarning("Не хватает библиотек", msg)


# ══════════════════════════════════════════════
#  Вкладка Яндекс Музыки
# ══════════════════════════════════════════════
class YandexTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._tracks = []
        self._build()

    def _build(self):
        pad = dict(padx=24, pady=6)

        # Токен
        tk.Label(self, text="Токен Яндекс Музыки", bg=BG2, fg=FG2,
                 font=FONT_MAIN).pack(anchor="w", **pad)
        row = ttk.Frame(self)
        row.pack(fill="x", padx=24, pady=(0,10))
        self.token_var = tk.StringVar()
        self.token_entry = ttk.Entry(row, textvariable=self.token_var,
                                     show="•", width=52)
        self.token_entry.pack(side="left", ipady=6, fill="x", expand=True)
        tk.Button(row, text="👁", bg=BG3, fg=FG2, relief="flat", bd=0,
                  command=self._toggle_token).pack(side="left", padx=(6,0), ipady=4)

        tk.Label(self, text="ℹ  Токен: music.yandex.ru → F12 → Application → Cookies → Session_id\n"
                             "  или используй yandex-music-token из GitHub",
                 bg=BG2, fg=FG2, font=("Segoe UI", 9), justify="left").pack(anchor="w", padx=24)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=24, pady=12)

        # Кнопки
        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", padx=24, pady=4)
        ttk.Button(btn_row, text="⬇  Загрузить треки из Яндекса",
                   style="Accent.TButton",
                   command=self._fetch).pack(side="left")
        self.save_fmt = tk.StringVar(value="xlsx")
        tk.Label(btn_row, text="  Формат:", bg=BG2, fg=FG2).pack(side="left")
        for fmt in ("xlsx", "txt", "csv"):
            tk.Radiobutton(btn_row, text=fmt, variable=self.save_fmt, value=fmt,
                           bg=BG2, fg=FG, selectcolor=BG3,
                           activebackground=BG2, font=FONT_MAIN).pack(side="left", padx=3)
        ttk.Button(btn_row, text="💾  Сохранить файл",
                   style="Red.TButton",
                   command=self._save).pack(side="right")

        # Прогресс
        self.ym_status = tk.StringVar(value="Ожидание…")
        tk.Label(self, textvariable=self.ym_status,
                 bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", padx=24, pady=(8,2))
        self.ym_bar = ttk.Progressbar(self, mode="determinate")
        self.ym_bar.pack(fill="x", padx=24, pady=(0,8))

        # Лог
        tk.Label(self, text="Лог", bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", padx=24)
        self.log = LogBox(self)
        self.log.pack(fill="both", expand=True, padx=24, pady=(4,16))

    def _toggle_token(self):
        cur = self.token_entry.cget("show")
        self.token_entry.configure(show="" if cur == "•" else "•")

    def _fetch(self):
        token = self.token_var.get().strip()
        if not token:
            messagebox.showwarning("Токен", "Введи токен Яндекс Музыки")
            return
        if "yandex_music" in missing:
            messagebox.showerror("Ошибка", "Установи: pip install yandex_music")
            return
        threading.Thread(target=self._fetch_thread, args=(token,), daemon=True).start()

    def _fetch_thread(self, token):
        self._tracks = []
        self.ym_status.set("Подключение к Яндекс Музыке…")
        self.log.clear()
        try:
            client = YMClient(token).init()
            self.log.append("✅ Подключено к Яндекс Музыке", SUCCESS)
            self.ym_status.set("Получение списка треков…")
            likes = client.users_likes_tracks()
            raw = likes.fetch_tracks()
            total = len(raw)
            self.ym_bar.configure(maximum=total)
            for i, track in enumerate(raw):
                artists = ", ".join(track.artists_name()) if track.artists_name() else "Unknown"
                self._tracks.append((artists, track.title or ""))
                self.ym_bar["value"] = i + 1
                if (i+1) % 50 == 0 or i+1 == total:
                    self.ym_status.set(f"Загружено {i+1} / {total}")
            self.log.append(f"🎵 Итого треков: {total}", SUCCESS)
            self.ym_status.set(f"Готово! Загружено {total} треков. Нажми «Сохранить файл»")
        except Exception as e:
            self.log.append(f"❌ Ошибка: {e}", ERROR)
            self.ym_status.set("Ошибка подключения")

    def _save(self):
        if not self._tracks:
            messagebox.showwarning("Нет данных", "Сначала загрузи треки из Яндекса")
            return
        fmt = self.save_fmt.get()
        ext = {"xlsx": ".xlsx", "txt": ".txt", "csv": ".csv"}[fmt]
        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=[(fmt.upper(), f"*{ext}"), ("All files", "*.*")],
            initialfile=f"tracks{ext}")
        if not path:
            return
        try:
            if fmt == "xlsx":
                self._save_xlsx(path)
            elif fmt == "txt":
                self._save_txt(path)
            else:
                self._save_csv(path)
            self.log.append(f"💾 Сохранено: {path}", SUCCESS)
            messagebox.showinfo("Готово", f"Файл сохранён:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _save_xlsx(self, path):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tracks"
        ws.append(["artist", "title"])
        for a, t in self._tracks:
            ws.append([a, t])
        wb.save(path)

    def _save_txt(self, path):
        with open(path, "w", encoding="utf-8") as f:
            for a, t in self._tracks:
                f.write(f"{a} — {t}\n")

    def _save_csv(self, path):
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["artist", "title"])
            for a, t in self._tracks:
                w.writerow([a, t])


# ══════════════════════════════════════════════
#  Вкладка YouTube Music
# ══════════════════════════════════════════════
class YouTubeTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._running  = False
        self._tracks   = []
        self._progress_file = "progress.txt"
        self._build()

    def _build(self):
        pad = dict(padx=24, pady=4)

        # browser.json
        tk.Label(self, text="Файл авторизации (browser.json)",
                 bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", **pad)
        r1 = ttk.Frame(self); r1.pack(fill="x", padx=24, pady=(0,8))
        self.auth_var = tk.StringVar(value="browser.json")
        ttk.Entry(r1, textvariable=self.auth_var).pack(side="left", ipady=6, fill="x", expand=True)
        ttk.Button(r1, text="📂", style="Accent.TButton",
                   command=lambda: self._browse(self.auth_var, [("JSON","*.json")])).pack(side="left", padx=(6,0))

        # Файл треков
        tk.Label(self, text="Файл треков (xlsx / txt / csv)",
                 bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", **pad)
        r2 = ttk.Frame(self); r2.pack(fill="x", padx=24, pady=(0,8))
        self.file_var = tk.StringVar()
        ttk.Entry(r2, textvariable=self.file_var).pack(side="left", ipady=6, fill="x", expand=True)
        ttk.Button(r2, text="📂", style="Accent.TButton",
                   command=lambda: self._browse(
                       self.file_var,
                       [("Все форматы","*.xlsx *.txt *.csv"),
                        ("Excel","*.xlsx"),("Text","*.txt"),("CSV","*.csv")]
                   )).pack(side="left", padx=(6,0))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=24, pady=10)

        # Кнопки управления
        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", padx=24, pady=4)
        self.start_btn = ttk.Button(btn_row, text="▶  Начать перенос",
                                    style="Accent.TButton", command=self._start)
        self.start_btn.pack(side="left")
        self.stop_btn = ttk.Button(btn_row, text="⏹  Стоп",
                                   style="Red.TButton", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=8)
        ttk.Button(btn_row, text="🗑  Сбросить прогресс",
                   command=self._reset_progress).pack(side="right")

        # Статус
        self.yt_status = tk.StringVar(value="Ожидание…")
        tk.Label(self, textvariable=self.yt_status,
                 bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", padx=24, pady=(10,2))

        # Прогресс-бар
        bar_row = ttk.Frame(self); bar_row.pack(fill="x", padx=24, pady=(0,6))
        self.yt_bar = ttk.Progressbar(bar_row, mode="determinate")
        self.yt_bar.pack(fill="x", expand=True, side="left")
        self.pct_var = tk.StringVar(value="0%")
        tk.Label(bar_row, textvariable=self.pct_var,
                 bg=BG2, fg=FG2, font=FONT_MAIN, width=5).pack(side="left", padx=(8,0))

        # Лог
        tk.Label(self, text="Лог", bg=BG2, fg=FG2, font=FONT_MAIN).pack(anchor="w", padx=24)
        self.log = LogBox(self)
        self.log.pack(fill="both", expand=True, padx=24, pady=(4,16))

    def _browse(self, var, ftypes):
        path = filedialog.askopenfilename(filetypes=ftypes)
        if path:
            var.set(path)

    # ── Загрузка треков из файла ───────────────
    def _load_tracks(self, path):
        ext = os.path.splitext(path)[1].lower()
        tracks = []
        if ext == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            # Пропускаем заголовок если есть
            start = 1 if rows and str(rows[0][0]).lower() in ("artist","артист") else 0
            for row in rows[start:]:
                if row and len(row) >= 2 and row[0] and row[1]:
                    tracks.append((str(row[0]), str(row[1])))
        elif ext == ".txt":
            with open(path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if " — " in line:
                        a, t = line.split(" — ", 1)
                        tracks.append((a.strip(), t.strip()))
                    elif line:
                        tracks.append(("", line))
        else:  # csv — поддерживаем оба формата (новый чистый и старый с [list])
            with open(path, encoding="utf-8-sig") as f:
                next(f, None)  # пропускаем заголовок
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    if line.startswith("["):
                        # Старый формат: ['artist'],title
                        try:
                            bracket_end = line.index("],")
                            artists_raw = line[:bracket_end+1]
                            title = line[bracket_end+2:]
                            artists = re.findall(r"'([^']+)'", artists_raw)
                            tracks.append((", ".join(artists), title))
                        except ValueError:
                            pass
                    else:
                        import csv
                        row = next(csv.reader([line]))
                        if len(row) >= 2:
                            tracks.append((row[0], row[1]))
        return tracks

    # ── Загрузка прогресса ────────────────────
    def _load_done(self):
        try:
            with open(self._progress_file, "r", encoding="utf-8") as f:
                return set(l.strip() for l in f if l.strip())
        except FileNotFoundError:
            return set()

    def _reset_progress(self):
        if os.path.exists(self._progress_file):
            if messagebox.askyesno("Сброс прогресса",
                                   "Удалить файл прогресса?\nПеренос начнётся с начала."):
                os.remove(self._progress_file)
                self.log.append("🗑 Прогресс сброшен", WARNING)
        else:
            messagebox.showinfo("Прогресс", "Файл прогресса не найден — перенос начнётся с начала")

    # ── Старт/Стоп ────────────────────────────
    def _start(self):
        auth  = self.auth_var.get().strip()
        fpath = self.file_var.get().strip()
        if not auth or not os.path.exists(auth):
            messagebox.showwarning("Авторизация", f"Файл не найден:\n{auth}")
            return
        if not fpath or not os.path.exists(fpath):
            messagebox.showwarning("Файл треков", "Выбери файл с треками")
            return
        if "ytmusicapi" in missing:
            messagebox.showerror("Ошибка", "Установи: pip install ytmusicapi")
            return
        self._running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.log.clear()
        threading.Thread(target=self._transfer_thread,
                         args=(auth, fpath), daemon=True).start()

    def _stop(self):
        self._running = False
        self.yt_status.set("⏹ Остановлено пользователем")
        self.stop_btn.configure(state="disabled")
        self.start_btn.configure(state="normal")

    # ── Поток переноса ─────────────────────────
    def _transfer_thread(self, auth, fpath):
        try:
            yt = YTMusic(auth)
            self.log.append("✅ Авторизация YouTube Music OK", SUCCESS)
        except Exception as e:
            self.log.append(f"❌ Ошибка авторизации: {e}", ERROR)
            self._done_ui()
            return

        try:
            self._tracks = self._load_tracks(fpath)
            self.log.append(f"📂 Загружено треков: {len(self._tracks)}", FG)
        except Exception as e:
            self.log.append(f"❌ Ошибка чтения файла: {e}", ERROR)
            self._done_ui()
            return

        done = self._load_done()
        skipped = sum(1 for a, t in self._tracks if f"{a} — {t}" in done)
        if skipped:
            self.log.append(f"⏩ Пропускаем уже перенесённые: {skipped}", WARNING)

        total    = len(self._tracks)
        ok = fail = skip = 0
        self.yt_bar.configure(maximum=total)

        with open(self._progress_file, "a", encoding="utf-8") as pf:
            for i, (artist, title) in enumerate(self._tracks):
                if not self._running:
                    break

                key   = f"{artist} — {title}"
                query = f"{artist} {title}"

                self.yt_bar["value"] = i + 1
                pct = int((i+1)/total*100)
                self.pct_var.set(f"{pct}%")
                self.yt_status.set(f"{i+1}/{total}  ·  ✅{ok}  ❌{fail}  ⏩{skip}")

                if key in done:
                    skip += 1
                    continue

                try:
                    results = yt.search(query, filter="songs")
                    if results:
                        yt.rate_song(results[0]["videoId"], "LIKE")
                        self.log.append(f"✅ {key}", SUCCESS)
                        ok += 1
                    else:
                        self.log.append(f"❌ Не найдено: {key}", ERROR)
                        fail += 1
                    pf.write(key + "\n")
                    pf.flush()
                    done.add(key)
                except Exception as e:
                    err = str(e)
                    if "401" in err or "Unauthorized" in err:
                        self.log.append(
                            f"🔑 Куки устарели! Обнови browser.json и запусти снова.\n"
                            f"   Остановлено на: {key}", WARNING)
                        self._running = False
                        break
                    else:
                        self.log.append(f"⚠️  {key}: {err}", WARNING)
                        fail += 1

        self.log.append(
            f"\n─── Итог ───\n✅ Добавлено: {ok}   ❌ Не найдено: {fail}   ⏩ Пропущено: {skip}",
            ACCENT)
        self.yt_status.set(f"Готово! ✅{ok}  ❌{fail}  ⏩{skip}")
        self._done_ui()

    def _done_ui(self):
        self._running = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")


# ══════════════════════════════════════════════
#  Виджет лога
# ══════════════════════════════════════════════
class LogBox(tk.Text):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG3, fg=FG, font=("Consolas", 9),
                         relief="flat", bd=0, state="disabled",
                         wrap="word", height=12, **kw)
        sb = ttk.Scrollbar(parent, command=self.yview)
        self.configure(yscrollcommand=sb.set)
        # Цветовые теги
        for tag, color in [("ok", SUCCESS), ("err", ERROR),
                            ("warn", WARNING), ("acc", ACCENT)]:
            self.tag_configure(tag, foreground=color)

    def append(self, text, color=None):
        tag_map = {SUCCESS: "ok", ERROR: "err", WARNING: "warn", ACCENT: "acc"}
        tag = tag_map.get(color)
        self.configure(state="normal")
        self.insert("end", text + "\n", tag or "")
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
