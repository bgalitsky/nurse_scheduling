"""Имитация сервера Ollama для тестов (настоящая модель в тестовой среде не запускается).

Отвечает как «идеальная» модель по словарю известных фрагментов + добавляет одну
выдуманную сущность, чтобы проверить, что инструмент отбрасывает то, чего нет в тексте.
Запуск вручную:  python tests/mock_ollama.py 11500
"""
import json
import re
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

DIGEST = "a" * 64
MODEL = "qwen3:8b"

KNOWN = [
    # синтетические PDF
    ("Мазуренко Галина Викторовна", "фио_пациента", ""), ("Мазуренко Г.В.", "фио_пациента", ""),
    ("Галины Викторовны", "фио_пациента", ""), ("Мазуренко", "фио_пациента", ""),
    ("Мазуренко Ольга Петровна", "фио_родственника", ""), ("Мазуренко Сергея Петровича", "фио_родственника", ""),
    ("Сидоров И.П.", "фио_врача", "кардиолог"), ("Сидоров", "фио_врача", "кардиолог"),
    ("Петрова А.А.", "фио_врача", ""), ("Кузнецов Андрей Михайлович", "фио_врача", ""),
    ("Лебедева Н.С.", "фио_врача", "невролог"), ("Орлова Е.В.", "фио_врача", ""),
    ("125009, г. Москва, ул. Тверская, д. 7, кв. 15", "адрес", ""),
    ("+7 (916) 123-45-67", "телефон", ""), ("8-903-555-12-34", "телефон", ""),
    ("g.mazurenko@mail.ru", "email", ""), ("123-456-789 01", "снилс", ""),
    ("7700 1234 5678 9012", "полис", ""), ("45 12 345678, выдан ОВД Тверской 12.03.2010", "паспорт", ""),
    ("12345/26", "номер_истории", ""), ("88", "номер_документа", ""),
    ("Иванов Пётр Сергеевич", "фио_пациента", ""), ("Иванова Мария Ивановна", "фио_родственника", ""),
    ("Волков Д.А.", "фио_врача", ""), ("9876", "номер_истории", ""),
    ("Московская обл., г. Химки, ул. Ленина, д. 3, кв. 8", "адрес", ""),
    # test1.docx
    ("Врачев Игорь Николаевич", "фио_пациента", ""), ("4150320877000093", "полис", ""),
    ("Москва, Люблинская улица, д. 72, к. 3, кв. 187", "адрес", ""),
    ("049-002-2-2025-0007427", "номер_документа", ""), ("2266K25000807-1", "номер_документа", ""),
    # проверка check-model
    ("Тестова Мария Олеговна", "фио_пациента", ""), ("+7 900 000-00-00", "телефон", ""),
    ("г. Казань, ул. Пробная, д. 1, кв. 2", "адрес", ""), ("77", "номер_документа", ""),
    ("Проверкин А.А.", "фио_врача", "кардиолог"), ("Тестова О.И.", "фио_родственника", ""),
]
HALLUCINATION = {"текст": "Несуществующий Иван Иванович", "категория": "фио_врача", "специальность": ""}
CALLS = {"chat": 0}


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, obj):
        b = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(b))); self.end_headers(); self.wfile.write(b)

    def do_GET(self):
        if self.path == "/api/tags":
            self._send({"models": [{"name": MODEL, "model": MODEL, "digest": DIGEST}]})
        elif self.path == "/api/version":
            self._send({"version": "mock"})
        else:
            self.send_error(404)

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        assert body["options"]["temperature"] == 0 and body["think"] is False and body["format"]
        CALLS["chat"] += 1
        chunk = body["messages"][-1]["content"]
        ents = [{"текст": t, "категория": c, "специальность": s} for t, c, s in KNOWN if t in chunk]
        ents.append(HALLUCINATION)
        self._send({"message": {"role": "assistant",
                                "content": json.dumps({"сущности": ents}, ensure_ascii=False)}})


def start(port=0):
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


if __name__ == "__main__":
    s = start(int(sys.argv[1]) if len(sys.argv) > 1 else 11500)
    print("mock ollama on", s.server_address)
    threading.Event().wait()
