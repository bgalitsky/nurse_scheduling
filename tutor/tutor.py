# app.py
# Streamlit Dialogue Math Tutor (kid-friendly) + levels + word problems + show-work + progress save + optional GPT support
#
# Run:
#   pip install streamlit openai
#   streamlit run app.py
#
# Optional GPT:
#   export OPENAI_API_KEY="..."
#
# Notes:
# - The app ALWAYS checks arithmetic problems deterministically.
# - GPT (if enabled) is used for conversational tutoring: explanations, hints, gentle coaching,
#   and (optionally) grading open-ended/word-problem answers when deterministic grading is hard.

import json
import math
import os
import random
import re
from configparser import RawConfigParser
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import streamlit as st

# GPT support (OpenAI Responses API)
try:
    from openai import OpenAI  # official SDK
except Exception:
    OpenAI = None

DEFAULT_MODEL = "gpt-5"

config_path='config.ini'
config = RawConfigParser()
config.read(config_path)

# Facebook credentials
email = config.get('Facebook', 'email')
password = config.get('Facebook', 'password')

# API keys
gemini_api_key = config.get('Gemini', 'api_key')
os.environ["OPENAI_API_KEY"] = config.get('OpenAI', 'api_key')

# -----------------------------
# Storage (simple local JSON)
# -----------------------------

PROGRESS_FILE = "math_tutor_progress.json"


def load_progress() -> Dict:
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_progress(data: Dict) -> None:
    try:
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        # If running in a read-only environment, fail gracefully
        pass


# -----------------------------
# Parsing problems from uploaded text
# -----------------------------

def parse_problems_from_text(text: str) -> List[str]:
    """
    Accepts .txt with:
      - one problem per line, OR
      - problems separated by blank lines, OR
      - numbered/bulleted lists
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return []

    # Split by blank lines
    chunks = [c.strip() for c in re.split(r"\n\s*\n", text) if c.strip()]

    if len(chunks) == 1:
        # Might be one-per-line or numbered list
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        numbered = sum(1 for ln in lines if re.match(r"^(\d+[\).\:]|[-•])\s+", ln))
        if numbered >= max(2, len(lines) // 3):
            probs = []
            for ln in lines:
                ln = re.sub(r"^(\d+[\).\:]|[-•])\s+", "", ln).strip()
                if ln:
                    probs.append(ln)
            return probs
        # If many lines, treat each line as a problem
        if len(lines) >= 2:
            return lines

    return chunks


# -----------------------------
# Deterministic math checking helpers
# -----------------------------

def safe_float(s: str) -> Optional[float]:
    s = s.strip().replace(",", ".")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        try:
            return float(s)
        except Exception:
            return None
    return None


def normalize_answer_text(s: str) -> str:
    return re.sub(r"\s+", "", s.strip().lower())


def extract_math_expression(problem: str) -> Optional[str]:
    """
    Very lightweight extraction: finds first "a op b" or "a op b op c".
    Also normalizes × ÷.
    """
    p = problem.lower().replace("×", "*").replace("÷", "/")

    m2 = re.search(r"(-?\d+)\s*([\+\-\*/])\s*(-?\d+)\s*([\+\-\*/])\s*(-?\d+)", p)
    if m2:
        a, op1, b, op2, c = m2.groups()
        return f"{a}{op1}{b}{op2}{c}"

    m = re.search(r"(-?\d+)\s*([\+\-\*/])\s*(-?\d+)", p)
    if m:
        a, op, b = m.group(1), m.group(2), m.group(3)
        return f"{a}{op}{b}"

    return None


def eval_expression(expr: str) -> Optional[float]:
    """
    Safe eval for basic arithmetic only.
    Allowed chars: digits + - * / ( ) . whitespace
    """
    if not re.fullmatch(r"[\d\+\-\*/\(\)\.\s]+", expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))
    except Exception:
        return None


def compute_correct_answer(problem: str) -> Tuple[Optional[str], Optional[float]]:
    expr = extract_math_expression(problem)
    if not expr:
        return (None, None)
    val = eval_expression(expr)
    if val is None:
        return (None, None)
    if abs(val - round(val)) < 1e-9:
        ival = int(round(val))
        return (str(ival), float(ival))
    return (str(val), float(val))


def is_correct_numeric(student: str, correct_value: float) -> bool:
    sf = safe_float(student)
    if sf is None:
        return False
    return abs(sf - correct_value) < 1e-9


def try_extract_number_from_text(s: str) -> Optional[float]:
    """
    For word problems, student might answer "It is 24 apples."
    Extract first number.
    """
    s = s.replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    return safe_float(m.group(0))


# -----------------------------
# Curriculum + Levels
# -----------------------------

@dataclass
class Concept:
    title: str
    explanation: str
    hint: str


CURRICULUM_BY_LEVEL: Dict[str, List[Concept]] = {
    "Grade 4 (easy)": [
        Concept(
            "Addition & subtraction (2-digit)",
            "Use place value: split into tens and ones. Add/subtract tens, then ones. If subtracting and the ones are too small, borrow 10 from the tens.",
            "Split into tens and ones. Do tens first, then ones."
        ),
        Concept(
            "Multiplication as equal groups",
            "Multiplication means repeated addition. 6×4 = six groups of 4. Use known facts like 5×something and adjust.",
            "Think: equal groups or repeated addition."
        ),
        Concept(
            "Division as sharing",
            "Division is equal sharing. 18÷3 asks what number times 3 gives 18.",
            "Ask: what times the divisor equals the total?"
        ),
    ],
    "Grade 5 (medium)": [
        Concept(
            "Multi-step arithmetic",
            "Do operations in order. Use parentheses if needed. For mental math: break numbers apart and recombine.",
            "Try doing it in smaller parts."
        ),
        Concept(
            "Fractions (simple)",
            "A fraction is part of a whole. 1/2 means one out of two equal parts. To find 1/4 of 20, divide 20 by 4.",
            "To find a fraction of a number: divide, then multiply."
        ),
        Concept(
            "Word problems (translate to math)",
            "Read carefully, underline numbers, decide which operation matches the story, then compute.",
            "What is being asked? Which operation fits the story?"
        ),
    ],
    "Grade 6 (challenge)": [
        Concept(
            "Ratios and rates",
            "A ratio compares two quantities. A rate is a ratio with different units (like km per hour).",
            "Write the ratio as a fraction and simplify."
        ),
        Concept(
            "Percent basics",
            "Percent means 'out of 100'. 25% of 40 = 0.25×40 = 10.",
            "Convert percent to a fraction or decimal."
        ),
        Concept(
            "Two-step equations (intro)",
            "Solve by undoing operations step-by-step. If x+5=12, subtract 5 from both sides: x=7.",
            "Do the opposite operation to both sides."
        ),
    ],
}

TOPICS = [
    "Mixed",
    "Addition/Subtraction",
    "Multiplication",
    "Division",
    "Fractions",
    "Word Problems",
    "Percent",
    "Ratios",
    "Equations",
]


# -----------------------------
# Problem generation (includes word problems)
# -----------------------------

def gen_arithmetic_problem(level: str, topic: str) -> Tuple[str, Optional[float], str]:
    """
    Returns (problem_text, correct_numeric_value_or_None, solution_explanation)
    """
    rng = random.Random()

    def explain(expr: str, answer: str, extra: str = "") -> str:
        return f"Expression: {expr}\nAnswer: {answer}\n{extra}".strip()

    if topic == "Mixed":
        topic = rng.choice(["Addition/Subtraction", "Multiplication", "Division", "Word Problems"])

    if topic == "Addition/Subtraction":
        if level == "Grade 4 (easy)":
            a, b = rng.randint(10, 99), rng.randint(10, 99)
        else:
            a, b = rng.randint(100, 999), rng.randint(10, 999)
        op = rng.choice(["+", "-"])
        if op == "-" and b > a:
            a, b = b, a
        expr = f"{a} {op} {b}"
        val = eval_expression(expr.replace(" ", ""))
        ans = str(int(val)) if val is not None else "?"
        problem = f"{expr} = ?"
        sol = explain(expr, ans, "Tip: use place value (hundreds/tens/ones).")
        return problem, float(int(val)) if val is not None else None, sol

    if topic == "Multiplication":
        if level == "Grade 4 (easy)":
            a, b = rng.randint(2, 9), rng.randint(2, 9)
        elif level == "Grade 5 (medium)":
            a, b = rng.randint(2, 12), rng.randint(2, 12)
        else:
            a, b = rng.randint(10, 30), rng.randint(2, 12)
        expr = f"{a} * {b}"
        val = eval_expression(expr)
        problem = f"{a} × {b} = ?"
        sol = explain(f"{a}×{b}", str(int(val)), "Think: equal groups or known facts.")
        return problem, float(int(val)), sol

    if topic == "Division":
        if level == "Grade 4 (easy)":
            b = rng.randint(2, 9)
            q = rng.randint(2, 12)
        elif level == "Grade 5 (medium)":
            b = rng.randint(2, 12)
            q = rng.randint(2, 20)
        else:
            b = rng.randint(2, 20)
            q = rng.randint(2, 25)
        a = b * q
        expr = f"{a} / {b}"
        val = eval_expression(expr)
        problem = f"{a} ÷ {b} = ?"
        sol = explain(f"{a}÷{b}", str(int(val)), "Ask: what times the divisor equals the total?")
        return problem, float(int(val)), sol

    if topic == "Fractions":
        # simple "fraction of a number"
        denom = random.choice([2, 3, 4, 5, 6, 8, 10])
        num = random.choice([1, 2, 3]) if denom >= 4 else 1
        if num >= denom:
            num = 1
        base = random.choice([12, 18, 20, 24, 30, 36, 40, 48, 60])
        # ensure divisible by denom for kid-friendly
        base = base - (base % denom)
        val = (num / denom) * base
        problem = f"What is {num}/{denom} of {base}?"
        sol = (
            f"To find {num}/{denom} of {base}:\n"
            f"1) Divide {base} by {denom}: {base} ÷ {denom} = {base//denom}\n"
            f"2) Multiply by {num}: {base//denom} × {num} = {int(val)}"
        )
        return problem, float(int(val)), sol

    if topic == "Percent":
        pct = random.choice([10, 20, 25, 30, 40, 50, 75])
        base = random.choice([20, 40, 60, 80, 100, 120, 200])
        val = (pct / 100) * base
        problem = f"What is {pct}% of {base}?"
        sol = f"{pct}% = {pct}/100. So {pct}/100 × {base} = {val}."
        return problem, float(val), sol

    if topic == "Ratios":
        a = random.randint(2, 12)
        b = random.randint(2, 12)
        g = math.gcd(a, b)
        problem = f"Simplify the ratio {a}:{b}."
        sol = f"Divide both parts by gcd({a},{b})={g}. So {a}:{b} = {a//g}:{b//g}."
        # ratio answer is non-numeric text; we will grade by normalized text comparison
        return problem, None, sol

    if topic == "Equations":
        # x + k = n or ax = n
        if random.random() < 0.5:
            x = random.randint(2, 20)
            k = random.randint(2, 20)
            n = x + k
            problem = f"Solve for x: x + {k} = {n}"
            sol = f"Subtract {k} from both sides: x = {n} - {k} = {x}."
            return problem, float(x), sol
        else:
            x = random.randint(2, 20)
            a = random.randint(2, 12)
            n = a * x
            problem = f"Solve for x: {a}x = {n}"
            sol = f"Divide both sides by {a}: x = {n} ÷ {a} = {x}."
            return problem, float(x), sol

    # Word Problems (default)
    # We produce a deterministic numeric answer.
    return gen_word_problem(level)


def gen_word_problem(level: str) -> Tuple[str, float, str]:
    rng = random.Random()
    templates = []

    # Grade-adjusted numbers
    if level == "Grade 4 (easy)":
        A = rng.randint(10, 60)
        B = rng.randint(5, 40)
        m = rng.randint(2, 9)
        n = rng.randint(2, 9)
    elif level == "Grade 5 (medium)":
        A = rng.randint(30, 150)
        B = rng.randint(10, 120)
        m = rng.randint(3, 12)
        n = rng.randint(3, 12)
    else:
        A = rng.randint(80, 400)
        B = rng.randint(30, 300)
        m = rng.randint(4, 20)
        n = rng.randint(4, 20)

    # Add/sub story
    templates.append((
        f"Maya has {A} stickers. She gives {B} stickers to a friend. How many stickers does she have left?",
        float(A - B),
        f"Subtract because she gives some away: {A} − {B} = {A - B}."
    ))

    # Multiplication story
    templates.append((
        f"There are {m} boxes. Each box has {n} oranges. How many oranges are there in total?",
        float(m * n),
        f"Multiply because it’s equal groups: {m} × {n} = {m * n}."
    ))

    # Division story (ensure divisible)
    total = m * n
    templates.append((
        f"{total} cookies are shared equally among {m} kids. How many cookies does each kid get?",
        float(n),
        f"Divide because it’s sharing equally: {total} ÷ {m} = {n}."
    ))

    problem, val, sol = rng.choice(templates)
    return problem, val, sol


def infer_topic(problem: str) -> str:
    p = problem.lower().replace("×", "*").replace("÷", "/")
    if "ratio" in p and ":" in p:
        return "Ratios"
    if "solve for x" in p or re.search(r"\bx\b", p):
        if re.search(r"\d+x", p) or re.search(r"\bx\s*\+\s*\d+", p):
            return "Equations"
    if "%" in p or "percent" in p:
        return "Percent"
    if re.search(r"\b\d+\s*/\s*\d+\b", p) and "of" in p:
        return "Fractions"
    if "*" in p:
        return "Multiplication"
    if "/" in p:
        return "Division"
    if "+" in p or "-" in p:
        return "Addition/Subtraction"
    # heuristic: story-like
    story_words = ["has", "gives", "left", "each", "total", "shared", "among", "how many"]
    if any(w in p for w in story_words):
        return "Word Problems"
    return "Mixed"


# -----------------------------
# Show-work validation (step checking)
# -----------------------------

def step_is_valid_equation(step: str) -> bool:
    """
    Accepts steps like:
      "34+27=61"
      "34 + 27 = 30+20 + 4+7"
      "30+20=50"
    We check if it contains exactly one '=' and both sides are valid arithmetic expressions.
    """
    s = step.strip().replace("×", "*").replace("÷", "/")
    if s.count("=") != 1:
        return False
    left, right = [t.strip() for t in s.split("=", 1)]
    if not left or not right:
        return False
    lv = eval_expression(left)
    rv = eval_expression(right)
    if lv is None or rv is None:
        return False
    return abs(lv - rv) < 1e-9


# -----------------------------
# GPT helper
# -----------------------------

def gpt_available() -> bool:
    return OpenAI is not None and bool(os.environ.get("OPENAI_API_KEY", "").strip())


def call_gpt_tutor(messages: List[Dict[str, str]], model: str) -> str:
    """
    Uses Responses API with conversation state (alternating role messages).
    """
    if not gpt_available():
        return "GPT is not available (missing OPENAI_API_KEY or openai package)."

    client = OpenAI()
    # Keep it concise / kid-friendly
    response = client.responses.create(
        model=model,
        input=messages,
    )
    return getattr(response, "output_text", "") or ""


def build_gpt_system_prompt(level: str, mode: str) -> str:
    return (
        "You are a friendly, patient math tutor for a 10–12 year old.\n"
        f"Target level: {level}.\n"
        "Rules:\n"
        "- Be encouraging, short, and clear.\n"
        "- Prefer hints before giving full answers.\n"
        "- If the student is stuck twice, give the full solution with steps.\n"
        "- Use simple language and show one step at a time.\n"
        "- Never claim you verified something if you didn't; rely on the app's provided correct answer when available.\n"
        f"Teaching style: {mode}.\n"
    )


# -----------------------------
# Dialogue engine state
# -----------------------------

def init_state():
    ss = st.session_state
    if "messages" not in ss:
        ss.messages = []  # chat transcript for UI
    if "tutor_memory" not in ss:
        ss.tutor_memory = []  # separate msg list for GPT conversation state
    if "problem_bank" not in ss:
        ss.problem_bank = []  # user-uploaded problems
    if "use_uploaded" not in ss:
        ss.use_uploaded = False
    if "awaiting_answer" not in ss:
        ss.awaiting_answer = False
    if "current_problem" not in ss:
        ss.current_problem = None
    if "current_answer_value" not in ss:
        ss.current_answer_value = None
    if "current_answer_text" not in ss:
        ss.current_answer_text = None
    if "current_solution" not in ss:
        ss.current_solution = None
    if "attempts" not in ss:
        ss.attempts = 0
    if "stuck_count" not in ss:
        ss.stuck_count = 0
    if "show_work_mode" not in ss:
        ss.show_work_mode = False
    if "pending_steps" not in ss:
        ss.pending_steps = []  # list of step strings
    if "practice_queue" not in ss:
        ss.practice_queue = []  # generated problems queue
    if "student_name" not in ss:
        ss.student_name = "Student"
    if "level" not in ss:
        ss.level = "Grade 4 (easy)"
    if "topic" not in ss:
        ss.topic = "Mixed"
    if "mode" not in ss:
        ss.mode = "Explain concept then practice"
    if "gpt_enabled" not in ss:
        ss.gpt_enabled = False
    if "gpt_model" not in ss:
        ss.gpt_model = "gpt-5.2"
    if "progress" not in ss:
        ss.progress = load_progress()


def tutor_say(text: str):
    st.session_state.messages.append({"role": "assistant", "content": text})


def student_say(text: str):
    st.session_state.messages.append({"role": "user", "content": text})


def progress_key() -> str:
    # per-student tracking
    return normalize_answer_text(st.session_state.student_name) or "student"


def record_result(correct: bool, problem: str, student_answer: str):
    data = st.session_state.progress
    key = progress_key()
    if key not in data:
        data[key] = {
            "student_name": st.session_state.student_name,
            "history": [],
            "stats": {"attempted": 0, "correct": 0},
        }
    data[key]["stats"]["attempted"] += 1
    if correct:
        data[key]["stats"]["correct"] += 1
    data[key]["history"].append({
        "ts": datetime.utcnow().isoformat() + "Z",
        "level": st.session_state.level,
        "topic": st.session_state.topic,
        "problem": problem,
        "student_answer": student_answer,
        "correct": correct,
    })
    save_progress(data)


def explain_concept_for_problem(problem: str) -> str:
    topic = infer_topic(problem)
    concepts = CURRICULUM_BY_LEVEL.get(st.session_state.level, CURRICULUM_BY_LEVEL["Grade 4 (easy)"])
    # pick closest concept by name heuristic
    if topic in ["Addition/Subtraction"]:
        return concepts[0].explanation
    if topic in ["Multiplication"]:
        return concepts[1].explanation if len(concepts) > 1 else concepts[0].explanation
    if topic in ["Division"]:
        return concepts[2].explanation if len(concepts) > 2 else concepts[0].explanation
    # fallback
    return concepts[-1].explanation


def hint_for_problem(problem: str) -> str:
    topic = infer_topic(problem)
    concepts = CURRICULUM_BY_LEVEL.get(st.session_state.level, CURRICULUM_BY_LEVEL["Grade 4 (easy)"])
    if topic == "Addition/Subtraction":
        return concepts[0].hint
    if topic == "Multiplication":
        return concepts[1].hint if len(concepts) > 1 else concepts[0].hint
    if topic == "Division":
        return concepts[2].hint if len(concepts) > 2 else concepts[0].hint
    return concepts[-1].hint


def next_problem_from_uploaded() -> Optional[str]:
    ss = st.session_state
    if not ss.problem_bank:
        return None
    # pop from front
    return ss.problem_bank.pop(0)


def next_problem_generated() -> str:
    ss = st.session_state
    # Keep a small queue so "next" is instant
    if not ss.practice_queue:
        for _ in range(5):
            p, val, sol = gen_arithmetic_problem(ss.level, ss.topic)
            ss.practice_queue.append((p, val, sol))
    p, val, sol = ss.practice_queue.pop(0)
    ss.current_solution = sol
    ss.current_answer_value = val
    # for arithmetic-like problems embedded in text, keep answer_text too
    if val is not None and abs(val - round(val)) < 1e-9:
        ss.current_answer_text = str(int(round(val)))
    else:
        ss.current_answer_text = str(val) if val is not None else None
    return p


def start_turn():
    ss = st.session_state
    ss.attempts = 0
    ss.stuck_count = 0
    ss.pending_steps = []

    problem = None
    val = None
    sol = None

    if ss.use_uploaded:
        problem = next_problem_from_uploaded()
        if problem is not None:
            topic = infer_topic(problem)
            # If we can compute deterministically (arithmetic), do it
            ct, cv = compute_correct_answer(problem)
            if cv is not None:
                val = cv
                sol = f"Computed from expression in the problem.\nAnswer: {ct}"
            ss.current_answer_value = val
            ss.current_answer_text = ct
            ss.current_solution = sol
    if problem is None:
        problem = next_problem_generated()

    ss.current_problem = problem
    ss.awaiting_answer = True

    if ss.mode == "Explain concept then practice":
        tutor_say(f"📚 Mini-lesson:\n\n{explain_concept_for_problem(problem)}")

    tutor_say(
        f"🧩 **Problem:** {problem}\n\n"
        "Type your answer. You can also type **hint** or **stuck**.\n"
        + ("If you want, turn on **Show your work** and enter steps like `34+27=61`.\n" if ss.show_work_mode else "")
    )


def gpt_coach(user_text: str, context: Dict) -> Optional[str]:
    """
    Returns a GPT message if enabled, else None.
    We pass the current problem + correct answer (if known) + hint + solution.
    """
    ss = st.session_state
    if not ss.gpt_enabled or not gpt_available():
        return None

    sys_prompt = build_gpt_system_prompt(ss.level, ss.mode)
    problem = context.get("problem", "")
    correct = context.get("correct", None)
    hint = context.get("hint", "")
    solution = context.get("solution", "")

    # Build a small, controlled state:
    gpt_msgs = [
        {"role": "system", "content": sys_prompt},
        {"role": "user", "content": f"Current problem: {problem}"},
    ]

    if correct is not None:
        gpt_msgs.append({"role": "user", "content": f"Correct answer (for you to rely on): {correct}"})
    if hint:
        gpt_msgs.append({"role": "user", "content": f"Hint you may use: {hint}"})
    if solution:
        gpt_msgs.append({"role": "user", "content": f"Solution steps (if needed): {solution}"})

    gpt_msgs.append({"role": "user", "content": f"Student said: {user_text}\nRespond as the tutor."})

    return call_gpt_tutor(gpt_msgs, ss.gpt_model).strip() or None


def grade_answer(student_text: str) -> Tuple[bool, str]:
    """
    Deterministic grading where possible.
    Returns (is_correct, feedback_string).
    """
    ss = st.session_state
    problem = ss.current_problem or ""
    val = ss.current_answer_value
    correct_text = ss.current_answer_text

    # Ratio text grading
    if infer_topic(problem) == "Ratios":
        # Expect simplified like "2:3" etc from solution
        # We'll compute it if possible
        m = re.search(r"ratio\s+(\d+)\s*:\s*(\d+)", problem.lower())
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            g = math.gcd(a, b)
            expected = f"{a//g}:{b//g}"
            ok = normalize_answer_text(student_text) == normalize_answer_text(expected)
            return ok, f"Expected: {expected}"

    # If we have numeric correct value, use it
    if val is not None:
        # For word problems, extract first number if needed
        sf = safe_float(student_text)
        if sf is None:
            extracted = try_extract_number_from_text(student_text)
            if extracted is not None:
                sf = extracted
        if sf is None:
            return False, "Please answer with a number (you can add words too, like '24 apples')."
        ok = abs(sf - val) < 1e-9
        return ok, f"Correct answer: {correct_text}"

    # Try computing from expression if present
    ct, cv = compute_correct_answer(problem)
    if cv is not None:
        ok = is_correct_numeric(student_text, cv)
        return ok, f"Correct answer: {ct}"

    # Otherwise can't deterministically grade
    return False, "I can't automatically grade that one. Explain your steps and I’ll help!"


def handle_show_work_steps(steps_text: str) -> Tuple[bool, List[str]]:
    """
    Validate each non-empty line as an equation that is true.
    """
    steps = [ln.strip() for ln in steps_text.split("\n") if ln.strip()]
    bad = []
    for s in steps:
        if not step_is_valid_equation(s):
            bad.append(s)
    return (len(bad) == 0), bad


def handle_student_input(user_text: str):
    ss = st.session_state
    student_say(user_text)

    if not ss.awaiting_answer:
        tutor_say("Type **start** to begin, or **next** for another problem.")
        return

    low = user_text.strip().lower()
    if low in {"hint", "clue", "help", "help me"}:
        ss.stuck_count += 1
        tutor_say(f"💡 Hint: {hint_for_problem(ss.current_problem)}")
        g = gpt_coach(user_text, {
            "problem": ss.current_problem,
            "correct": ss.current_answer_text,
            "hint": hint_for_problem(ss.current_problem),
            "solution": ss.current_solution
        })
        if g:
            tutor_say(g)
        return

    if low in {"stuck", "i'm stuck", "i am stuck", "dont know", "don't know", "i don't know", "idk", "no idea"}:
        ss.stuck_count += 1
        if ss.stuck_count == 1:
            tutor_say(f"That’s okay. Here’s a smaller hint: {hint_for_problem(ss.current_problem)}\n\nTry again.")
            g = gpt_coach(user_text, {
                "problem": ss.current_problem,
                "correct": ss.current_answer_text,
                "hint": hint_for_problem(ss.current_problem),
                "solution": ss.current_solution
            })
            if g:
                tutor_say(g)
            return
        # second time stuck -> full solution
        tutor_say("No worries — let’s do it together step-by-step:")
        if ss.current_answer_text is not None:
            tutor_say(f"✅ **Correct answer:** {ss.current_answer_text}")
        if ss.current_solution:
            tutor_say(f"🧠 Steps:\n\n```\n{ss.current_solution}\n```")
        g = gpt_coach(user_text, {
            "problem": ss.current_problem,
            "correct": ss.current_answer_text,
            "hint": hint_for_problem(ss.current_problem),
            "solution": ss.current_solution
        })
        if g:
            tutor_say(g)
        record_result(False, ss.current_problem, user_text)
        ss.awaiting_answer = False
        tutor_say("Want another one? Type **next**.")
        return

    # Optional show-work validation (if enabled and student provided steps via dedicated UI)
    # The main answer is still graded from user_text.
    ss.attempts += 1
    ok, feedback = grade_answer(user_text)

    if ok:
        tutor_say(f"✅ Correct! {feedback}")
        record_result(True, ss.current_problem, user_text)
        ss.awaiting_answer = False
        # Add GPT praise + micro-explanation (optional)
        g = gpt_coach(user_text, {
            "problem": ss.current_problem,
            "correct": ss.current_answer_text,
            "hint": hint_for_problem(ss.current_problem),
            "solution": ss.current_solution
        })
        if g:
            tutor_say(g)
        tutor_say("Type **next** for another problem.")
        return

    # incorrect
    if ss.attempts == 1:
        tutor_say(f"Not quite. 💡 Hint: {hint_for_problem(ss.current_problem)}")
        g = gpt_coach(user_text, {
            "problem": ss.current_problem,
            "correct": ss.current_answer_text,
            "hint": hint_for_problem(ss.current_problem),
            "solution": ss.current_solution
        })
        if g:
            tutor_say(g)
        return

    # second attempt wrong -> show full correction
    tutor_say("Thanks for trying. Here’s the correct solution:")
    if ss.current_answer_text is not None:
        tutor_say(f"✅ **Correct answer:** {ss.current_answer_text}")
    if ss.current_solution:
        tutor_say(f"🧠 Steps:\n\n```\n{ss.current_solution}\n```")

    g = gpt_coach(user_text, {
        "problem": ss.current_problem,
        "correct": ss.current_answer_text,
        "hint": hint_for_problem(ss.current_problem),
        "solution": ss.current_solution
    })
    if g:
        tutor_say(g)

    record_result(False, ss.current_problem, user_text)
    ss.awaiting_answer = False
    tutor_say("Ready for another? Type **next**.")


# -----------------------------
# UI
# -----------------------------

st.set_page_config(page_title="Dialogue Math Tutor + GPT", page_icon="🧠", layout="centered")
init_state()
ss = st.session_state

st.title("🧠 Dialogue Math Tutor (with Levels + Show-Work + GPT option)")
st.caption("Dialogue-first tutoring. Deterministic grading for arithmetic, optional GPT for coaching and open-ended feedback.")

with st.sidebar:
    st.header("Student")
    ss.student_name = st.text_input("Student name", value=ss.student_name)

    st.markdown("---")
    st.header("Curriculum")
    ss.level = st.selectbox("Difficulty / Level", list(CURRICULUM_BY_LEVEL.keys()), index=list(CURRICULUM_BY_LEVEL.keys()).index(ss.level))
    ss.topic = st.selectbox("Topic", TOPICS, index=TOPICS.index(ss.topic))
    ss.mode = st.radio("Flow", ["Explain concept then practice", "Practice only"], index=0 if ss.mode == "Explain concept then practice" else 1)

    st.markdown("---")
    st.header("Show your work")
    ss.show_work_mode = st.toggle("Enable step-by-step checking", value=ss.show_work_mode)
    if ss.show_work_mode:
        st.caption("Enter steps like: `34+27=61` or `30+20=50`. Each line must be a TRUE equation.")

    st.markdown("---")
    st.header("Upload problems (.txt)")
    uploaded = st.file_uploader("Upload a .txt with one problem per line or separated by blank lines.", type=["txt"])
    if uploaded is not None:
        txt = uploaded.read().decode("utf-8", errors="replace")
        probs = parse_problems_from_text(txt)
        if probs:
            ss.problem_bank = probs
            ss.use_uploaded = True
            st.success(f"Loaded {len(probs)} uploaded problem(s).")
        else:
            st.warning("No problems detected. Try one problem per line or separate by blank lines.")

    ss.use_uploaded = st.toggle("Use uploaded problems first", value=ss.use_uploaded)

    st.markdown("---")
    st.header("GPT conversational support")
    if not gpt_available():
        st.info("GPT is available if you install `openai` and set `OPENAI_API_KEY`.")
    ss.gpt_enabled = st.toggle("Enable GPT tutor replies", value=ss.gpt_enabled and gpt_available(), disabled=not gpt_available())
    ss.gpt_model = st.text_input("Model", value=ss.gpt_model, help="Example: gpt-5.2")

    st.markdown("---")
    if st.button("🔄 Reset chat (keep progress)"):
        ss.messages = []
        ss.awaiting_answer = False
        ss.current_problem = None
        ss.current_answer_value = None
        ss.current_answer_text = None
        ss.current_solution = None
        ss.attempts = 0
        ss.stuck_count = 0
        tutor_say(f"Hi {ss.student_name}! Type **start** when you’re ready.")

    if st.button("🧹 Clear uploaded problem bank"):
        ss.problem_bank = []
        ss.use_uploaded = False
        st.success("Cleared uploaded problems.")

    st.markdown("---")
    st.header("Progress")
    key = progress_key()
    pdata = ss.progress.get(key, {}).get("stats", {"attempted": 0, "correct": 0})
    attempted = pdata.get("attempted", 0)
    correct = pdata.get("correct", 0)
    acc = (correct / attempted * 100.0) if attempted else 0.0
    st.write(f"Attempted: **{attempted}**")
    st.write(f"Correct: **{correct}**")
    st.write(f"Accuracy: **{acc:.1f}%**")

# Seed greeting
# Seed greeting
if not ss.messages:
    tutor_say(f"Hi {ss.student_name}! I’m your math tutor. Type **start** to begin.")

# --- INPUT FIRST (so new messages appear on the same run) ---
user_text = st.chat_input("Type: start / next / hint / stuck / or your answer...")

if user_text:
    low = user_text.strip().lower()

    if low in {"start", "begin"}:
        student_say(user_text)
        tutor_say("Great — let’s begin!")
        start_turn()

    elif low in {"next", "another", "more"}:
        student_say(user_text)
        start_turn()

    else:
        handle_student_input(user_text)

    # Force immediate refresh so the just-added message is visible right away
    st.rerun()

# --- THEN RENDER CHAT HISTORY ---
for msg in ss.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Show-work panel (optional) AFTER rendering is fine
if ss.show_work_mode and ss.awaiting_answer:
    with st.expander("✍️ Show your work (optional)", expanded=False):
        show_work_text = st.text_area(
            "Steps (one per line)",
            height=120,
            placeholder="Example:\n34+27=61\n30+20=50\n4+7=11\n50+11=61"
        )
        if st.button("Check my steps"):
            if show_work_text.strip():
                ok, bad = handle_show_work_steps(show_work_text)
                if ok:
                    st.success("All steps look mathematically correct. Nice!")
                else:
                    st.error("Some steps are not valid TRUE equations:")
                    st.write(bad)
            else:
                st.info("Enter at least one step.")
