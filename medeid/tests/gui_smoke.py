"""Сквозная проверка окна приложения на вымышленных документах (без настоящей модели).

    xvfb-run -a python tests/gui_smoke.py <папка_программы> <временный_HOME> [папка_для_снимков]

Папка настроек создаётся в <временный_HOME>/settings (MEDEID_SETTINGS), исходные документы
берутся из <временный_HOME>/data_in.
"""
import os
import sys
import time

proj, home = sys.argv[1], sys.argv[2]
shots = sys.argv[3] if len(sys.argv) > 3 else None
os.environ["HOME"] = home
os.environ["MEDEID_SETTINGS"] = os.path.join(home, "settings")
sys.path.insert(0, proj); sys.path.insert(0, os.path.join(proj, "tests"))

import yaml  # noqa: E402
import mock_ollama  # noqa: E402
from medeid.settings import init_settings  # noqa: E402

srv = mock_ollama.start()
sd = init_settings(None, "tester")
cfg_path = os.path.join(sd, "config.yaml")
base = yaml.safe_load(open(cfg_path, encoding="utf-8"))
base["модель"]["сервер"] = f"http://127.0.0.1:{srv.server_address[1]}"
base["модель"]["дайджест"] = mock_ollama.DIGEST
base["категории_и_режимы"]["фио_пациента"] = "сквозной_код"
base["сквозной_код_пациента"] = "да"
open(cfg_path, "w", encoding="utf-8").write(yaml.safe_dump(base, allow_unicode=True, sort_keys=False))
open(os.path.join(sd, "roles.yaml"), "w", encoding="utf-8").write("операторы: [tester]\nадминистраторы: [admin1]\n")

import tkinter as tk  # noqa: E402
from tkinter import filedialog, messagebox  # noqa: E402

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


assert app.roles_for("tester") == ["оператор"] and app.roles_for("admin1") == ["администратор"]
root = tk.Tk(); root.geometry("1100x780")
w = app.OperatorWindow(root, "tester")
w.v_in.set(os.path.join(home, "data_in"))
w.v_order.set("142-о"); w.v_date.set("01.11.2026"); w.v_dept.set("онкология")
w.v_llm.set(True); w._llm_toggle()
w._run()
pump(root, lambda: any(s[1] in ("Партия обработана", "medeid", "Конфигурация") for s in shown))
print("партия:", shown[-1][2].replace("\n", " | "))
shot("operator_window.png")
w.tv_batches.selection_set(w.tv_batches.get_children()[0])
w._verify(); pump(root, lambda: any(s[1] == "Повторяемость" for s in shown), 120)
print("повторяемость:", [s[2] for s in shown if s[1] == "Повторяемость"])
root.destroy()

root = tk.Tk(); root.geometry("1100x780")
a = app.AdminWindow(root, "admin1")
kids = a.tv_runs.get_children(); a.tv_runs.selection_set(kids[-1]); a._show_run()
saved = os.path.join(home, "rules_out.docx")
filedialog.asksaveasfilename = lambda **k: (shown.append(("save", k.get("initialfile"), "")), saved)[1]
a._save_rules("docx")
root.update(); time.sleep(0.3); root.update()
shot("admin_window.png")
print("администратор: партий", len(kids), "документов", len(a.tv_files.get_children()),
      "ключей", len(a.tv_keys.get_children()), "| правила:", a.lbl_rules.cget("text"),
      "| файл по умолчанию:", [s[1] for s in shown if s[0] == "save"], "сохранён:", os.path.exists(saved))
root.destroy()
srv.shutdown()
