"""
app.py — Desktop GUI for food schedule generator.
Cross-platform: macOS and Windows.
Run: python app.py
"""
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from rozkladka_core import (
    detect_dates_from_filename,
    generate_daily,
    convert_xlsx_to_pdf,
    scale_nakladna,
)

_FONT       = ('Arial', 10)
_FONT_SMALL = ('Arial', 9)
_FONT_LOG   = ('Courier New', 9)
_PAD        = {'padx': 6, 'pady': 4}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Генератор розкладки')
        self.minsize(600, 500)
        self.resizable(True, True)

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=8)

        nb.add(DailyTab(nb),  text='  По днях  ')
        nb.add(PeriodTab(nb), text='  Зведена  ')
        nb.add(ScaleTab(nb),  text='  Накладна  ')


# ── Shared base ───────────────────────────────────────────────────────────────

class _BaseTab(ttk.Frame):

    def _row(self, parent, row, label, var, btn_label, btn_cmd):
        """Label + Entry + Button row."""
        ttk.Label(parent, text=label, font=_FONT).grid(row=row, column=0, sticky='e', **_PAD)
        e = ttk.Entry(parent, textvariable=var, width=44, font=_FONT)
        e.grid(row=row, column=1, sticky='ew', **_PAD)
        ttk.Button(parent, text=btn_label, command=btn_cmd).grid(row=row, column=2, **_PAD)

    def _make_log(self, parent, row):
        """Scrollable dark log widget spanning 3 columns."""
        frame = ttk.LabelFrame(parent, text='Лог')
        frame.grid(row=row, column=0, columnspan=3, sticky='nsew', padx=6, pady=6)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        txt = tk.Text(frame, height=10, font=_FONT_LOG, wrap='word', state='disabled',
                      background='#1e1e1e', foreground='#d4d4d4', relief='flat',
                      insertbackground='white')
        sb = ttk.Scrollbar(frame, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.grid(row=0, column=0, sticky='nsew')
        sb.grid(row=0, column=1, sticky='ns')
        return txt

    def _log(self, txt, msg):
        """Thread-safe append to log widget."""
        def _do():
            txt.configure(state='normal')
            txt.insert('end', msg + '\n')
            txt.see('end')
            txt.configure(state='disabled')
        self.after(0, _do)

    def _clear_log(self, txt):
        txt.configure(state='normal')
        txt.delete('1.0', 'end')
        txt.configure(state='disabled')

    def _autofill_date(self, path_var, date_var):
        """Try to parse week start date from filename and fill date_var."""
        path = path_var.get()
        if path:
            detected = detect_dates_from_filename(path)
            if detected:
                date_var.set(detected)


# ── Tab 1: По днях ────────────────────────────────────────────────────────────

class DailyTab(_BaseTab):
    def __init__(self, parent):
        super().__init__(parent)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

        self._xlsx     = tk.StringVar()
        self._out_file = tk.StringVar()
        self._unit     = tk.StringVar(value='Т0920')
        self._date     = tk.StringVar()

        self._row(self, 0, 'XLSX файл:', self._xlsx, 'Вибрати…', self._pick_xlsx)
        self._row(self, 1, 'Зберегти PDF як:', self._out_file, 'Зберегти…', self._pick_outfile)

        ttk.Label(self, text='Підрозділ:', font=_FONT).grid(row=2, column=0, sticky='e', **_PAD)
        ttk.Entry(self, textvariable=self._unit, width=20, font=_FONT).grid(
            row=2, column=1, sticky='w', **_PAD)

        ttk.Label(self, text='Дата початку тижня:', font=_FONT).grid(row=3, column=0, sticky='e', **_PAD)
        df = ttk.Frame(self)
        df.grid(row=3, column=1, sticky='w', **_PAD)
        ttk.Entry(df, textvariable=self._date, width=14, font=_FONT).pack(side='left')
        ttk.Label(df, text='ДД.ММ.РРРР', font=_FONT_SMALL, foreground='gray').pack(side='left', padx=6)

        self._btn = ttk.Button(self, text='Згенерувати PDF', command=self._run)
        self._btn.grid(row=4, column=0, columnspan=3, pady=8)

        self._log_txt = self._make_log(self, 5)

    def _pick_xlsx(self):
        p = filedialog.askopenfilename(
            title='Вибрати XLSX файл',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Всі файли', '*.*')])
        if p:
            self._xlsx.set(p)
            self._autofill_date(self._xlsx, self._date)
            if not self._out_file.get():
                date = self._date.get() or 'rozkladka'
                stem = f'rozkladka_{date.replace(".", "-")}'
                self._out_file.set(str(Path(p).parent / f'{stem}.pdf'))

    def _pick_outfile(self):
        date = self._date.get() or 'rozkladka'
        stem = f'rozkladka_{date.replace(".", "-")}'
        p = filedialog.asksaveasfilename(
            title='Зберегти PDF',
            defaultextension='.pdf',
            initialfile=f'{stem}.pdf',
            filetypes=[('PDF', '*.pdf'), ('Всі файли', '*.*')])
        if p:
            self._out_file.set(p)

    def _run(self):
        xlsx = self._xlsx.get().strip()
        out  = self._out_file.get().strip()
        unit = self._unit.get().strip() or 'Т0920'
        date = self._date.get().strip()

        if not xlsx:
            messagebox.showerror('Помилка', 'Вкажіть XLSX файл'); return
        if not out:
            messagebox.showerror('Помилка', 'Вкажіть шлях для збереження PDF'); return
        if not date:
            messagebox.showerror('Помилка', 'Вкажіть дату початку тижня (ДД.ММ.РРРР)'); return

        self._clear_log(self._log_txt)
        self._btn.configure(state='disabled')

        def _task():
            try:
                generate_daily(xlsx, out, unit, date,
                               progress_cb=lambda m: self._log(self._log_txt, m))

                self._log(self._log_txt, '\nГотово!')
            except Exception as e:
                self._log(self._log_txt, f'❌ Помилка: {e}')
            finally:
                self.after(0, lambda: self._btn.configure(state='normal'))

        threading.Thread(target=_task, daemon=True).start()


# ── Tab 2: Зведена ────────────────────────────────────────────────────────────

class PeriodTab(_BaseTab):
    def __init__(self, parent):
        super().__init__(parent)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(3, weight=1)

        self._xlsx     = tk.StringVar()
        self._out_file = tk.StringVar()

        self._row(self, 0, 'XLSX файл:', self._xlsx, 'Вибрати…', self._pick_xlsx)
        self._row(self, 1, 'Зберегти PDF як:', self._out_file, 'Зберегти…', self._pick_outfile)

        self._btn = ttk.Button(self, text='Конвертувати в PDF', command=self._run)
        self._btn.grid(row=2, column=0, columnspan=3, pady=8)

        self._log_txt = self._make_log(self, 3)

    def _pick_xlsx(self):
        p = filedialog.askopenfilename(
            title='Вибрати XLSX файл',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Всі файли', '*.*')])
        if p:
            self._xlsx.set(p)
            if not self._out_file.get():
                self._out_file.set(str(Path(p).with_suffix('.pdf')))

    def _pick_outfile(self):
        p = filedialog.asksaveasfilename(
            title='Зберегти PDF',
            defaultextension='.pdf',
            filetypes=[('PDF', '*.pdf'), ('Всі файли', '*.*')])
        if p:
            self._out_file.set(p)

    def _run(self):
        xlsx = self._xlsx.get().strip()
        out  = self._out_file.get().strip()

        if not xlsx:
            messagebox.showerror('Помилка', 'Вкажіть XLSX файл'); return
        if not out:
            messagebox.showerror('Помилка', 'Вкажіть вихідний файл'); return

        self._clear_log(self._log_txt)
        self._btn.configure(state='disabled')

        def _task():
            try:
                convert_xlsx_to_pdf(xlsx, out,
                                    progress_cb=lambda m: self._log(self._log_txt, m))
                self._log(self._log_txt, '\nГотово!')
            except Exception as e:
                self._log(self._log_txt, f'❌ Помилка: {e}')
            finally:
                self.after(0, lambda: self._btn.configure(state='normal'))

        threading.Thread(target=_task, daemon=True).start()


# ── Tab 3: Накладна ───────────────────────────────────────────────────────────

class ScaleTab(_BaseTab):
    def __init__(self, parent):
        super().__init__(parent)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

        self._src     = tk.StringVar()
        self._out_dir = tk.StringVar()
        self._base    = tk.IntVar(value=520)
        self._targets = tk.StringVar(value='70, 450')

        self._row(self, 0, 'Вхідна накладна:', self._src, 'Вибрати…', self._pick_src)
        self._row(self, 1, 'Вихідна папка:', self._out_dir, 'Вибрати…', self._pick_outdir)

        ttk.Label(self, text='База (осіб):', font=_FONT).grid(row=2, column=0, sticky='e', **_PAD)
        ttk.Spinbox(self, textvariable=self._base, from_=1, to=99999, width=9, font=_FONT).grid(
            row=2, column=1, sticky='w', **_PAD)

        ttk.Label(self, text='Цільові кількості:', font=_FONT).grid(row=3, column=0, sticky='e', **_PAD)
        tf = ttk.Frame(self)
        tf.grid(row=3, column=1, sticky='w', **_PAD)
        ttk.Entry(tf, textvariable=self._targets, width=20, font=_FONT).pack(side='left')
        ttk.Label(tf, text='(через кому)', font=_FONT_SMALL, foreground='gray').pack(side='left', padx=6)

        self._btn = ttk.Button(self, text='Масштабувати', command=self._run)
        self._btn.grid(row=4, column=0, columnspan=3, pady=8)

        self._log_txt = self._make_log(self, 5)

    def _pick_src(self):
        p = filedialog.askopenfilename(
            title='Вибрати накладну',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Всі файли', '*.*')])
        if p:
            self._src.set(p)
            if not self._out_dir.get():
                self._out_dir.set(str(Path(p).parent))

    def _pick_outdir(self):
        p = filedialog.askdirectory(title='Вибрати вихідну папку')
        if p:
            self._out_dir.set(p)

    def _run(self):
        src  = self._src.get().strip()
        out  = self._out_dir.get().strip()
        base = self._base.get()

        try:
            targets = [int(x.strip()) for x in self._targets.get().split(',') if x.strip()]
        except ValueError:
            messagebox.showerror('Помилка', 'Цільові кількості мають бути цілими числами через кому')
            return

        if not src:
            messagebox.showerror('Помилка', 'Вкажіть вхідну накладну'); return
        if not out:
            messagebox.showerror('Помилка', 'Вкажіть вихідну папку'); return
        if not targets:
            messagebox.showerror('Помилка', 'Вкажіть хоча б одну цільову кількість'); return

        self._clear_log(self._log_txt)
        self._btn.configure(state='disabled')

        def _task():
            try:
                scale_nakladna(src, out, base, targets,
                               progress_cb=lambda m: self._log(self._log_txt, m))
                self._log(self._log_txt, '\nГотово!')
            except Exception as e:
                self._log(self._log_txt, f'❌ Помилка: {e}')
            finally:
                self.after(0, lambda: self._btn.configure(state='normal'))

        threading.Thread(target=_task, daemon=True).start()


if __name__ == '__main__':
    app = App()
    app.mainloop()
