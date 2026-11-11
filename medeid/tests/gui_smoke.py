"""Сквозная проверка окна приложения на вымышленных документах (без настоящей модели).

    xvfb-run -a python tests/gui_smoke.py <копия_проекта> <временный_HOME> [папка_для_снимков]

Запускать на КОПИИ проекта: тест пишет config.yaml и roles.yaml в папку проекта.
"""
import os
import sys
import time

proj, home = sys.argv[1], sys.argv[2]
shots = sys.argv[3] if len(sys.argv) > 3 else None
os.environ["HOME"] = home
sys.path.insert(0, proj); sys.path.insert(0, os.path.join(proj, "tests"))

import yaml  # noqa: E402
import mock_ollama  # noqa: E402

srv = mock_ollama.start()
base = yaml.safe_load(open(os.path.join(proj, "config.example.yaml"), encoding="utf-8"))
base["модель"]["сервер"] = f"http://127.0.0.1:{srv.server_address[1]}"
base["модель"]["дайджест"] = mock_ollama.DIGEST
base["типы_документов"] = ["выписной_эпикриз", "протокол_консилиума", "осмотр_врача"]
open(os.path.join(proj, "config.yaml"), "w", encoding="utf-8").write(yaml.safe_dump(base, allow_unicode=True, sort_keys=False))

import tkinter as tk  # noqa: E402
from tkinter import messagebox  # noqa: E402

shown = []
messagebox.showinfo = lambda t, m: shown.append(("info", t, m))
messagebox.showerror = lambda t, m: shown.append(("error", t, m))
messagebox.showwarning = lambda t, m: shown.append(("warn", t, m))
messagebox.askyesno = lambda *a, **k: True
from medeid import app  # noqa: E402


def pump(root, cond, timeout=120):
    t0 = time.time()
    while not cond() and time.time() - t0 < timeout:
        root.update(); time.sleep(0.05)
    for _ in range(10):
        root.update(); time.sleep(0.05)


def shot(name):
    if shots:
        from PIL import ImageGrab
        ImageGrab.grab(xdisplay=os.environ["DISPLAY"]).crop((0, 0, 1100, 780)).save(os.path.join(shots, name))


root = tk.Tk(); root.geometry("1100x780")
w = app.OperatorWindow(root, "tester")
w.v_in.set(os.path.join(home, "src"))
w.v_order.set("142-о"); w.v_date.set("01.11.2026"); w.v_dept.set("онкология")
w.v_llm.set(True); w._llm_toggle()
w.v_modes["фио_врачей"].set("специальность")
w.v_code.set(True); w._code_toggle()
w._run()
pump(root, lambda: any(s[1] in ("Партия обработана", "medeid") for s in shown))
print("партия:", shown[-1][2].replace("\n", " | "))
shot("operator.png")
w.tv_batches.selection_set(w.tv_batches.get_children()[0])
w._verify(); pump(root, lambda: any(s[1] == "Повторяемость" for s in shown), 120)
print("повторяемость:", [s[2] for s in shown if s[1] == "Повторяемость"])
root.destroy()

root = tk.Tk(); root.geometry("1100x780")
a = app.AdminWindow(root, "admin1")
kids = a.tv_runs.get_children(); a.tv_runs.selection_set(kids[-1]); a._show_run()
root.update(); time.sleep(0.3); root.update()
shot("admin.png")
print("администратор: партий", len(kids), "документов", len(a.tv_files.get_children()),
      "ключей", len(a.tv_keys.get_children()))
root.destroy()
srv.shutdown()
