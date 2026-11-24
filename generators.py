"""Synthetic data generation. All values are fictional - never real personal data.

Uses Faker when installed; otherwise falls back to built-in word pools so the
application still works. Every built-in category and the custom dataset share
the same ID / seed / range / duplicate machinery.
"""
from __future__ import annotations

import datetime as dt
import random
import string
import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Optional

from config_manager import (ConfigError, column_range, normalize_column, opt_float,
                            parse_date, validate_columns)

try:  # Faker is the preferred source of realistic values
    from faker import Faker
    HAVE_FAKER = True
except Exception:  # pragma: no cover - depends on environment
    Faker = None  # type: ignore
    HAVE_FAKER = False


class GenerationAborted(Exception):
    """Raised inside generation/export when the app is closing or Force Stop is used."""


@dataclass
class ColMeta:
    name: str
    kind: str            # text | id | int | decimal | currency | date | bool
    decimals: int = 0


@dataclass
class Dataset:
    category: str
    cols: list
    rows: list
    generated_at: dt.datetime
    seed: Optional[int]
    locale: str
    engine_name: str

    @property
    def headers(self) -> list:
        return [c.name for c in self.cols]


# ------------------------------------------------------------------ word pools
FIRST_M = ("James Liam Noah Oliver Elijah Lucas Mateo Henry Daniel Samuel Joseph Carlos Miguel Andrew Ethan Ryan "
           "Nathan Adrian Victor Marco Hugo Felix Arjun Kenji Omar Tomas Gabriel Jonas Leo Ivan Rafael Diego "
           "Anton Dmitri Kofi Sven Pavel Mario Jun Rohan").split()
FIRST_F = ("Olivia Emma Sophia Isabella Mia Amelia Harper Evelyn Abigail Emily Ella Grace Chloe Lily Hannah Zoe "
           "Nora Camila Valentina Sofia Maria Elena Aisha Mei Yuki Priya Anika Freya Ingrid Lucia Carmen Rosa "
           "Katya Ayesha Leila Nadia Bianca Tessa Gwen Clara").split()
LAST = ("Smith Johnson Williams Brown Jones Garcia Miller Davis Rodriguez Martinez Hernandez Lopez Gonzalez Wilson "
        "Anderson Thomas Taylor Moore Jackson Martin Lee Perez Thompson White Harris Sanchez Clark Ramirez Lewis "
        "Robinson Walker Young Allen King Wright Scott Torres Nguyen Hill Flores Green Adams Nelson Baker Hall "
        "Rivera Campbell Mitchell Carter Roberts Santos Reyes Cruz Bautista Dela-Cruz Tanaka Suzuki Kim Park Chen "
        "Wang Singh Patel Kumar Khan Ali Hassan Novak Kowalski Ivanov Petrov Schmidt Muller Fischer Rossi Ferrari "
        "Bianchi Dubois Laurent Moreau Nilsson Larsen Jensen Okafor Mensah Abara").split()
STREETS = ("Maple Oak Cedar Pine Elm Washington Lake Hill Sunset Park Main Church Willow Highland Meadow River "
           "Forest Spring Ridge Franklin Jefferson Lincoln Madison Orchard Mill Bay Harbor Garden Cherry").split()
STREET_SFX = ["Street", "Avenue", "Road", "Lane", "Drive", "Boulevard", "Court", "Way", "Place", "Terrace"]
PLACES = [
    ("New York", "United States"), ("Los Angeles", "United States"), ("Chicago", "United States"),
    ("Houston", "United States"), ("Toronto", "Canada"), ("Vancouver", "Canada"), ("London", "United Kingdom"),
    ("Manchester", "United Kingdom"), ("Dublin", "Ireland"), ("Paris", "France"), ("Lyon", "France"),
    ("Berlin", "Germany"), ("Munich", "Germany"), ("Madrid", "Spain"), ("Barcelona", "Spain"),
    ("Rome", "Italy"), ("Milan", "Italy"), ("Amsterdam", "Netherlands"), ("Stockholm", "Sweden"),
    ("Oslo", "Norway"), ("Warsaw", "Poland"), ("Lisbon", "Portugal"), ("Sydney", "Australia"),
    ("Melbourne", "Australia"), ("Auckland", "New Zealand"), ("Tokyo", "Japan"), ("Osaka", "Japan"),
    ("Seoul", "South Korea"), ("Singapore", "Singapore"), ("Manila", "Philippines"), ("Cebu City", "Philippines"),
    ("Angeles City", "Philippines"), ("Bangkok", "Thailand"), ("Mumbai", "India"), ("Bengaluru", "India"),
    ("Dubai", "United Arab Emirates"), ("Cairo", "Egypt"), ("Nairobi", "Kenya"), ("Lagos", "Nigeria"),
    ("Sao Paulo", "Brazil"), ("Buenos Aires", "Argentina"), ("Mexico City", "Mexico"), ("Santiago", "Chile"),
]
COMPANY_SFX = ["Ltd", "Group", "Supply Co", "Trading", "Industries", "Partners", "Logistics", "Wholesale", "& Sons"]
LOREM = ("alpha amber anchor arrow atlas basil beacon birch bridge canyon cedar cobalt comet copper coral crest "
         "delta ember falcon fern field flint forge garnet glacier harbor hazel horizon indigo iris jade lagoon "
         "lantern lotus maple meadow mesa nectar nova oasis onyx orchid pebble pine prairie quartz raven reef "
         "ridge river sable sierra solstice summit tundra velvet willow zephyr").split()

ADJECTIVES = ("Premium Compact Wireless Ergonomic Classic Deluxe Portable Smart Heavy-Duty Eco Pro Ultra Essential "
              "Modern Rugged Lightweight Foldable Digital Vintage Advanced Everyday").split()
CATALOG = {  # category -> (nouns, (min cost, max cost))
    "Electronics": (["Headphones", "Keyboard", "Mouse", "Speaker", "Webcam", "Charger", "Monitor Stand", "Power Bank", "Router", "Smart Plug"], (12, 420)),
    "Home & Kitchen": (["Blender", "Kettle", "Cutting Board", "Storage Set", "Lamp", "Cookware Set", "Air Fryer", "Dinnerware Set", "Vacuum Flask", "Toaster"], (8, 180)),
    "Office Supplies": (["Notebook", "Desk Organizer", "Stapler", "Whiteboard", "Pen Set", "File Cabinet", "Label Maker", "Planner", "Paper Shredder", "Desk Pad"], (2, 140)),
    "Sports & Outdoors": (["Backpack", "Water Bottle", "Yoga Mat", "Tent", "Dumbbell Set", "Hiking Boots", "Sleeping Bag", "Cycling Helmet", "Jump Rope", "Camping Stove"], (6, 260)),
    "Clothing": (["Jacket", "T-Shirt", "Jeans", "Sneakers", "Hoodie", "Cap", "Scarf", "Gloves", "Socks Pack", "Raincoat"], (5, 150)),
    "Health & Beauty": (["Face Serum", "Shampoo", "Electric Toothbrush", "Hair Dryer", "Sunscreen", "Moisturizer", "Massage Gun", "Body Lotion", "Razor Set", "Lip Balm"], (3, 130)),
    "Toys & Games": (["Puzzle", "Board Game", "Building Blocks", "RC Car", "Plush Toy", "Art Kit", "Card Game", "Science Kit", "Doll House", "Drone"], (4, 110)),
}
PAYMENT_METHODS = ["Credit Card", "Debit Card", "PayPal", "Bank Transfer", "Cash", "Mobile Wallet", "Gift Card"]
PAYMENT_W = [34, 22, 14, 8, 10, 9, 3]
ORDER_STATUS = ["Completed", "Shipped", "Processing", "Pending", "Cancelled", "Refunded"]
ORDER_STATUS_W = [62, 14, 9, 7, 5, 3]
DEPARTMENTS = {
    "Engineering": [("Software Engineer", 70000, 130000), ("Senior Software Engineer", 110000, 170000), ("QA Engineer", 55000, 95000), ("DevOps Engineer", 80000, 140000), ("Engineering Manager", 130000, 190000)],
    "Sales": [("Sales Representative", 38000, 75000), ("Account Executive", 60000, 115000), ("Sales Manager", 85000, 150000)],
    "Marketing": [("Marketing Specialist", 45000, 80000), ("Content Strategist", 50000, 90000), ("Marketing Manager", 85000, 140000)],
    "Human Resources": [("HR Coordinator", 40000, 65000), ("Recruiter", 48000, 85000), ("HR Manager", 80000, 130000)],
    "Finance": [("Accountant", 50000, 90000), ("Financial Analyst", 60000, 110000), ("Finance Manager", 95000, 160000)],
    "Operations": [("Operations Associate", 35000, 60000), ("Logistics Coordinator", 40000, 70000), ("Operations Manager", 80000, 135000)],
    "Customer Support": [("Support Agent", 30000, 50000), ("Support Team Lead", 48000, 75000)],
    "Legal": [("Paralegal", 45000, 75000), ("Legal Counsel", 100000, 190000)],
    "IT": [("IT Technician", 38000, 68000), ("Systems Administrator", 62000, 105000), ("Security Analyst", 80000, 135000)],
}
EMP_STATUS = ["Active", "On Leave", "Probation", "Terminated"]
EMP_STATUS_W = [82, 5, 6, 7]
COURSES = ["Introduction to Computer Science", "Calculus I", "Linear Algebra", "Organic Chemistry", "World History",
           "Microeconomics", "Macroeconomics", "Statistics", "Psychology 101", "English Composition", "Physics I",
           "Data Structures", "Database Systems", "Marketing Principles", "Financial Accounting", "Business Ethics",
           "Graphic Design", "Environmental Science", "Sociology", "Philosophy", "Biology I", "Nursing Fundamentals",
           "Civil Engineering Basics", "Creative Writing"]
WAREHOUSES = ["Main Warehouse", "North Hub", "South Hub", "East Depot", "West Depot", "Central DC"]
FIN_TYPES = [  # (label, is_credit, weight, (min, max))
    ("Deposit", True, 22, (50, 5000)), ("Salary Credit", True, 7, (800, 6500)), ("Transfer In", True, 8, (20, 2500)),
    ("Refund", True, 4, (5, 400)), ("Interest", True, 3, (0.5, 120)), ("Withdrawal", False, 14, (20, 900)),
    ("Payment", False, 20, (10, 1600)), ("Transfer Out", False, 9, (20, 2500)), ("Fee", False, 3, (1, 45)),
    ("ATM Withdrawal", False, 10, (20, 400)),
]


def _ascii_slug(text: str) -> str:
    t = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return "".join(ch for ch in t.lower() if ch.isalnum())


def _years_ago(d: dt.date, years: int) -> dt.date:
    try:
        return d.replace(year=d.year - years)
    except ValueError:  # Feb 29
        return d.replace(year=d.year - years, day=28)


# ------------------------------------------------------------------ synthetic source
class Synth:
    """Faker-backed when available, otherwise built-in pools. Output is always fictional."""

    SAFE_DOMAINS = ("example.com", "example.org", "example.net")  # reserved by RFC 2606 - never real mailboxes

    def __init__(self, rng: random.Random, faker=None):
        self.rng = rng
        self.f = faker

    def _try(self, *names):
        if self.f is None:
            return None
        for n in names:
            fn = getattr(self.f, n, None)
            if fn is None:
                continue
            try:
                return fn()
            except Exception:
                continue
        return None

    def first_name(self, gender=None) -> str:
        if gender == "Male":
            v = self._try("first_name_male", "first_name")
        elif gender == "Female":
            v = self._try("first_name_female", "first_name")
        else:
            v = self._try("first_name")
        if v:
            return v
        pool = FIRST_M if gender == "Male" else FIRST_F if gender == "Female" else FIRST_M + FIRST_F
        return self.rng.choice(pool)

    def last_name(self) -> str:
        return self._try("last_name") or self.rng.choice(LAST)

    def full_name(self, gender=None) -> str:
        return f"{self.first_name(gender)} {self.last_name()}"

    def email(self, first: str, last: str) -> str:
        a, b = _ascii_slug(first), _ascii_slug(last)
        r = self.rng
        local = r.choice([f"{a}.{b}", f"{a}{b}", f"{a[:1]}{b}", f"{b}.{a}"]) if a and b else f"user{r.randint(1000, 99999)}"
        return f"{local}{r.randint(1, 999)}@{r.choice(self.SAFE_DOMAINS)}"

    def phone(self) -> str:
        v = self._try("phone_number")
        if v:
            return v
        return f"({self.rng.randint(201, 989)}) 555-01{self.rng.randint(0, 99):02d}"  # 555-01xx = fictional range

    def street(self) -> str:
        v = self._try("street_address")
        if v:
            return v.replace("\n", ", ")
        return f"{self.rng.randint(1, 9999)} {self.rng.choice(STREETS)} {self.rng.choice(STREET_SFX)}"

    def place(self) -> tuple:
        if self.f is not None:
            city = self._try("city")
            country = self._try("current_country", "country")
            if city and country:
                return city, country
        return self.rng.choice(PLACES)

    def full_address(self) -> str:
        city, country = self.place()
        return f"{self.street()}, {city}, {country}"

    def company(self) -> str:
        v = self._try("company")
        return v or f"{self.rng.choice(LAST)} {self.rng.choice(COMPANY_SFX)}"

    def words(self, n: int) -> str:
        if self.f is not None:
            try:
                return " ".join(self.f.words(nb=n)).capitalize()
            except Exception:
                pass
        return " ".join(self.rng.choice(LOREM) for _ in range(n)).capitalize()


# ------------------------------------------------------------------ ID factory
class IdFactory:
    def __init__(self, rng, s: dict, seq_state: dict, n: int):
        self.rng, self.s, self.seq, self.n = rng, s, seq_state, n

    def maker(self, key: str, default_prefix: str = "", width: int = 6, mode: str = "Global",
              prefix: str = "", suffix: str = "", kind: str = "digits") -> Callable:
        """Return fn(infix='') -> unique ID string. Prefix/suffix are applied around the core."""
        s = self.s
        mode = s["id_mode"] if mode == "Global" else mode
        pre = s["id_prefix"] or prefix or default_prefix
        suf = s["id_suffix"] or suffix
        rng = self.rng
        if mode == "Sequential":
            w = max(width, len(str(self.n + self.seq.get(key, 1))))

            def seq_id(infix=""):
                v = self.seq.get(key, 1)
                self.seq[key] = v + 1
                return f"{pre}{infix}{str(v).zfill(w)}{suf}"
            return seq_id
        w = max(width, len(str(self.n)) + 3)
        alphabet = string.digits if kind == "digits" else string.ascii_uppercase + string.digits
        used: set = set()

        def rand_id(infix=""):
            while True:
                core = "".join(rng.choices(alphabet, k=w))
                if kind == "digits" and core[0] == "0":
                    core = str(rng.randint(1, 9)) + core[1:]
                full = f"{pre}{infix}{core}{suf}"
                if full not in used:
                    used.add(full)
                    return full
        return rand_id


# ------------------------------------------------------------------ build context
class Ctx:
    def __init__(self, rng, syn, s, ids, n, d0, d1, progress, abort):
        self.rng, self.syn, self.s, self.ids, self.n = rng, syn, s, ids, n
        self.d0, self.d1 = d0, d1
        self.progress, self.abort = progress, abort
        self.step = max(1, n // 200)
        self.today = dt.date.today()
        self.last_name_parts = None   # lets a custom Email column follow the preceding Name column
        self.nmin, self.nmax = opt_float(s["num_min"]), opt_float(s["num_max"])

    def report(self, i: int) -> None:
        if self.abort is not None and self.abort.is_set():
            raise GenerationAborted()
        if self.progress:
            self.progress(min(1.0, i / max(1, self.n)))

    def money_range(self, lo: float, hi: float) -> tuple:
        """Category default range, overridden by the user's min/max if given."""
        a = self.nmin if self.nmin is not None else lo
        b = self.nmax if self.nmax is not None else hi
        if a > b:
            a, b = (a, a) if self.nmin is not None else (b, b)
        return a, b

    def rand_date(self, lo=None, hi=None) -> dt.date:
        lo, hi = lo or self.d0, hi or self.d1
        return dt.date.fromordinal(self.rng.randint(lo.toordinal(), hi.toordinal()))

    def sorted_dates(self, count: int) -> list:
        lo, hi = self.d0.toordinal(), self.d1.toordinal()
        return [dt.date.fromordinal(o) for o in sorted(self.rng.randint(lo, hi) for _ in range(count))]


# ------------------------------------------------------------------ category builders
PERSONAL_COLS = [ColMeta("Person ID", "id"), ColMeta("First Name", "text"), ColMeta("Last Name", "text"),
                 ColMeta("Age", "int"), ColMeta("Gender", "text"), ColMeta("Birthday", "date"),
                 ColMeta("Email", "text"), ColMeta("Phone", "text"), ColMeta("Address", "text"),
                 ColMeta("City", "text"), ColMeta("Country", "text")]


def build_personal(c: Ctx) -> list:
    rng, syn, s = c.rng, c.syn, c.s
    mk = c.ids.maker("personal:id", "P", 6)
    lo, hi = sorted((s["age_min"], s["age_max"]))
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        gender = rng.choices(["Male", "Female", "Non-binary"], [48, 48, 4])[0]
        first, last = syn.first_name(gender), syn.last_name()
        age = rng.randint(lo, hi)
        latest = _years_ago(c.today, age)
        earliest = _years_ago(c.today, age + 1) + dt.timedelta(days=1)
        bday = earliest + dt.timedelta(days=rng.randint(0, (latest - earliest).days))
        city, country = syn.place()
        rows.append([mk(), first, last, age, gender, bday, syn.email(first, last), syn.phone(), syn.street(), city, country])
    return rows


EMP_COLS = [ColMeta("Employee ID", "id"), ColMeta("Full Name", "text"), ColMeta("Department", "text"),
            ColMeta("Position", "text"), ColMeta("Salary", "currency", 2), ColMeta("Hire Date", "date"),
            ColMeta("Status", "text")]


def build_employees(c: Ctx) -> list:
    rng, syn = c.rng, c.syn
    mk = c.ids.maker("employees:id", "EMP-", 5)
    depts = list(DEPARTMENTS)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        dept = rng.choice(depts)
        pos, lo, hi = rng.choice(DEPARTMENTS[dept])
        lo, hi = c.money_range(lo, hi)
        salary = float(round(rng.uniform(lo, hi) / 100) * 100) if c.nmin is None and c.nmax is None else float(round(rng.uniform(lo, hi)))
        rows.append([mk(), syn.full_name(), dept, pos, salary, c.rand_date(),
                     rng.choices(EMP_STATUS, EMP_STATUS_W)[0]])
    return rows


SALES_COLS = [ColMeta("Transaction ID", "id"), ColMeta("Order Date", "date"), ColMeta("Customer", "text"),
              ColMeta("Product", "text"), ColMeta("Category", "text"), ColMeta("Quantity", "int"),
              ColMeta("Unit Price", "currency", 2), ColMeta("Total", "currency", 2),
              ColMeta("Payment Method", "text"), ColMeta("Order Status", "text")]


def build_sales(c: Ctx) -> list:
    rng, syn = c.rng, c.syn
    mk = c.ids.maker("sales:id", "TXN-", 7)
    customers = [syn.full_name() for _ in range(max(10, min(c.n, max(20, c.n // 4))))]
    cats = list(CATALOG)
    dates = c.sorted_dates(c.n)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        cat = rng.choice(cats)
        nouns, (plo, phi) = CATALOG[cat]
        lo, hi = c.money_range(plo * 1.3, phi * 1.8)
        price = round(rng.uniform(lo, hi), 2)
        qty = rng.choices([1, 2, 3, 4, 5, 6, 8, 10], [50, 25, 12, 6, 3, 2, 1, 1])[0]
        rows.append([mk(), dates[i], rng.choice(customers), f"{rng.choice(ADJECTIVES)} {rng.choice(nouns)}", cat, qty,
                     price, round(qty * price, 2), rng.choices(PAYMENT_METHODS, PAYMENT_W)[0],
                     rng.choices(ORDER_STATUS, ORDER_STATUS_W)[0]])
    return rows


FIN_COLS = [ColMeta("Account Number", "text"), ColMeta("Date", "date"), ColMeta("Transaction Type", "text"),
            ColMeta("Credit", "currency", 2), ColMeta("Debit", "currency", 2), ColMeta("Balance", "currency", 2),
            ColMeta("Reference", "id")]


def build_financial(c: Ctx) -> list:
    rng = c.rng
    mk = c.ids.maker("financial:ref", "REF-", 8, kind="alnum")
    k = max(1, min(200, c.n // 25))
    accounts, seen = [], set()
    while len(accounts) < k:
        a = "".join(rng.choices(string.digits, k=10))
        if a not in seen:
            seen.add(a)
            accounts.append(a)
    olo, ohi = c.money_range(500, 25000) if (c.nmin is not None or c.nmax is not None) else (500, 25000)
    balance = {a: round(rng.uniform(olo, ohi), 2) for a in accounts}
    labels = [t[0] for t in FIN_TYPES]
    weights = [t[2] for t in FIN_TYPES]
    dates = c.sorted_dates(c.n)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        acct = rng.choice(accounts)
        idx = rng.choices(range(len(FIN_TYPES)), weights)[0]
        label, is_credit, _, (alo, ahi) = FIN_TYPES[idx]
        if c.nmin is not None or c.nmax is not None:
            alo, ahi = c.money_range(alo, ahi)
        amt = round(rng.uniform(alo, ahi), 2)
        bal = balance[acct]
        if not is_credit and amt > bal:           # never overdraw: fall back to a deposit
            label, is_credit = "Deposit", True
            amt = round(rng.uniform(max(alo, 50), max(ahi, 100)), 2)
        bal = round(bal + amt if is_credit else bal - amt, 2)
        balance[acct] = bal
        rows.append([acct, dates[i], label, amt if is_credit else 0.0, 0.0 if is_credit else amt, bal, mk()])
    return rows


INV_COLS = [ColMeta("Product ID", "id"), ColMeta("SKU", "id"), ColMeta("Product Name", "text"),
            ColMeta("Category", "text"), ColMeta("Stock", "int"), ColMeta("Supplier", "text"),
            ColMeta("Unit Cost", "currency", 2), ColMeta("Unit Price", "currency", 2),
            ColMeta("Warehouse", "text"), ColMeta("Last Restocked", "date")]


def build_inventory(c: Ctx) -> list:
    rng, syn = c.rng, c.syn
    mk_id = c.ids.maker("inventory:id", "PRD-", 6)
    mk_sku = c.ids.maker("inventory:sku", "", 6, kind="alnum")
    suppliers = [syn.company() for _ in range(max(5, min(60, c.n // 20)))]
    cats = list(CATALOG)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        cat = rng.choice(cats)
        nouns, (clo, chi) = CATALOG[cat]
        markup = rng.uniform(1.25, 2.2)
        if c.nmin is not None or c.nmax is not None:
            plo, phi = c.money_range(clo * 1.25, chi * 2.2)
            price = round(rng.uniform(plo, phi), 2)
            cost = round(price / markup, 2)
        else:
            cost = round(rng.uniform(clo, chi), 2)
            price = round(cost * markup, 2)
        stock = 0 if rng.random() < 0.05 else int(rng.triangular(1, 800, 120))
        rows.append([mk_id(), mk_sku(cat[:3].upper() + "-"), f"{rng.choice(ADJECTIVES)} {rng.choice(nouns)}", cat, stock,
                     rng.choice(suppliers), cost, price, rng.choice(WAREHOUSES), c.rand_date()])
    return rows


EDU_COLS = [ColMeta("Student ID", "id"), ColMeta("Full Name", "text"), ColMeta("Course", "text"),
            ColMeta("Grade", "decimal", 1), ColMeta("Letter Grade", "text"), ColMeta("Enrollment Date", "date"),
            ColMeta("Attendance %", "decimal", 1)]


def _letter(g: float) -> str:
    for cut, ltr in ((93, "A"), (90, "A-"), (87, "B+"), (83, "B"), (80, "B-"), (77, "C+"), (73, "C"), (70, "C-"), (60, "D")):
        if g >= cut:
            return ltr
    return "F"


def build_education(c: Ctx) -> list:
    rng, syn = c.rng, c.syn
    mk = c.ids.maker("education:id", "STU-", 6)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        grade = round(min(100.0, max(35.0, rng.gauss(78, 12))), 1)
        att = round(min(100.0, max(50.0, rng.gauss(92, 7))), 1)
        rows.append([mk(), syn.full_name(), rng.choice(COURSES), grade, _letter(grade), c.rand_date(), att])
    return rows


BUILTIN = {
    "Personal Information": (PERSONAL_COLS, build_personal),
    "Employees": (EMP_COLS, build_employees),
    "Business/Sales": (SALES_COLS, build_sales),
    "Financial": (FIN_COLS, build_financial),
    "Inventory": (INV_COLS, build_inventory),
    "Education": (EDU_COLS, build_education),
}


# ------------------------------------------------------------------ custom dataset
def _custom_column(c: Ctx, col: dict):
    rng, syn = c.rng, c.syn
    t, dec = col["type"], col["decimals"]
    if t == "text":
        lo, hi = column_range(col)
        lo, hi = max(1, int(lo)), max(1, int(hi))
        return (lambda: syn.words(rng.randint(lo, hi))), ColMeta(col["name"], "text")
    if t == "integer":
        lo, hi = column_range(col)
        lo, hi = int(lo), int(hi)
        return (lambda: rng.randint(lo, hi)), ColMeta(col["name"], "int")
    if t in ("decimal", "currency"):
        lo, hi = column_range(col)
        d = 2 if t == "currency" else dec
        return (lambda: round(rng.uniform(lo, hi), d)), ColMeta(col["name"], t, d)
    if t == "name":
        def make_name():
            c.last_name_parts = (syn.first_name(), syn.last_name())
            return f"{c.last_name_parts[0]} {c.last_name_parts[1]}"
        return make_name, ColMeta(col["name"], "text")
    if t == "email":
        def make_email():
            f, l = c.last_name_parts or (syn.first_name(), syn.last_name())
            return syn.email(f, l)
        return make_email, ColMeta(col["name"], "text")
    if t == "phone":
        return syn.phone, ColMeta(col["name"], "text")
    if t == "address":
        return syn.full_address, ColMeta(col["name"], "text")
    if t == "date":
        lo = parse_date(col["min"], c.d0)
        hi = parse_date(col["max"], c.d1)
        return (lambda: c.rand_date(lo, hi)), ColMeta(col["name"], "date")
    if t == "boolean":
        p = col["true_pct"]
        return (lambda: rng.random() * 100 < p), ColMeta(col["name"], "bool")
    if t == "id":
        mk = c.ids.maker(f"custom:{col['name']}", "", col["width"], col["id_mode"], col["prefix"], col["suffix"], "digits")
        return mk, ColMeta(col["name"], "id")
    vals = col["values"]
    return (lambda: rng.choice(vals)), ColMeta(col["name"], "text")  # custom_list


def build_custom(c: Ctx, columns: list) -> tuple:
    cols = [normalize_column(x) for x in columns]
    errs = validate_columns(cols)
    if errs:
        raise ConfigError("; ".join(errs))
    makers, metas = [], []
    for col in cols:
        m, meta = _custom_column(c, col)
        makers.append(m)
        metas.append(meta)
    rows = []
    for i in range(c.n):
        if not i % c.step:
            c.report(i)
        rows.append([m() for m in makers])
    return metas, rows


# ------------------------------------------------------------------ public generator
class DataGenerator:
    """Stateful across cycles only for sequential-ID counters (`seq`)."""

    def __init__(self, seq_state: dict | None = None):
        self.seq: dict = dict(seq_state or {})
        self._fakers: dict = {}

    def reset_sequences(self) -> None:
        self.seq.clear()

    @property
    def engine_name(self) -> str:
        if HAVE_FAKER:
            import faker
            return f"Faker {getattr(faker, 'VERSION', '')}".strip()
        return "built-in generator (Faker not installed)"

    def _faker(self, locale: str):
        if not HAVE_FAKER:
            return None
        if locale not in self._fakers:
            try:
                self._fakers[locale] = Faker(locale)
            except Exception:
                self._fakers[locale] = Faker("en_US")
        return self._fakers[locale]

    def generate(self, category: str, settings: dict, custom_columns=None, records: int | None = None,
                 cycle: int = 0, progress=None, abort=None) -> Dataset:
        s = settings
        n = int(records if records is not None else s["records"])
        if n < 1:
            raise ConfigError("Records per file must be at least 1.")
        d0 = parse_date(s["date_start"], dt.date(2015, 1, 1))
        d1 = parse_date(s["date_end"], dt.date.today())
        if d0 > d1:
            raise ConfigError("Date range: start is after end.")
        if opt_float(s["num_min"]) is not None and opt_float(s["num_max"]) is not None \
                and opt_float(s["num_min"]) > opt_float(s["num_max"]):
            raise ConfigError("Number range: min is greater than max.")

        seed = int(s["seed"]) + cycle if s["use_seed"] else None   # fixed seed -> reproducible but different each cycle
        rng = random.Random(seed)
        faker = self._faker(s["locale"])
        if faker is not None:
            faker.seed_instance(rng.getrandbits(64))
        syn = Synth(rng, faker)

        dups = 0
        if s["allow_duplicates"] and n > 1:
            dups = min(n - 1, int(round(n * float(s["duplicate_pct"]) / 100.0)))
        base_n = n - dups
        ids = IdFactory(rng, s, self.seq, base_n)
        ctx = Ctx(rng, syn, s, ids, base_n, d0, d1, progress, abort)

        if category == "Custom Dataset":
            cols, rows = build_custom(ctx, custom_columns or [])
        elif category in BUILTIN:
            cols, builder = BUILTIN[category]
            rows = builder(ctx)
        else:
            raise ConfigError(f"Unknown category '{category}'.")

        if dups:
            copies: dict = {}
            for idx in (rng.randrange(base_n) for _ in range(dups)):
                copies[idx] = copies.get(idx, 0) + 1
            out = []
            for i, r in enumerate(rows):
                out.append(r)
                for _ in range(copies.get(i, 0)):
                    out.append(list(r))
            rows = out
        if progress:
            progress(1.0)
        return Dataset(category, list(cols), rows, dt.datetime.now(), seed, s["locale"], self.engine_name)
