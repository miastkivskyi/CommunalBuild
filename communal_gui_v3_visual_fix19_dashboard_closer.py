
import json
import os
from datetime import datetime
import ttkbootstrap as tb
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook

DATA_FILE = "meters.json"

TARIFFS = {
    "day": 4.32,
    "night": 2.16,
    "water": 16.0,
    "sewer": 11.0,
    "management": 419.0,
    "vat": 56.0,
    "intercom": 30.0,
}

def load_data():
    if not os.path.exists(DATA_FILE):
        return {"history": [], "tariffs": TARIFFS.copy()}
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "history" not in data:
        data["history"] = []
    if "tariffs" not in data:
        data["tariffs"] = TARIFFS.copy()
    return data

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

class App:
    def __init__(self):
        self.data = load_data()
        TARIFFS.update(self.data.get("tariffs", {}))

        self.root = tb.Window(themename="darkly")
        self.root.title("Комунальні платежі v3")
        self.root.geometry("1280x850")
        self.root.minsize(1200, 800)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_main = tb.Frame(notebook)
        self.tab_meters = tb.Frame(notebook)
        self.tab_costs = tb.Frame(notebook)
        self.tab_tariffs = tb.Frame(notebook)

        notebook.add(self.tab_main, text="🏠 Головна")
        notebook.add(self.tab_meters, text="📋 Показники")
        notebook.add(self.tab_costs, text="💰 Нарахування")
        notebook.add(self.tab_tariffs, text="⚙️ Тарифи")

        self.search_meters = tb.StringVar()
        self.search_costs = tb.StringVar()

        self.build_main()
        self.build_meters()
        self.build_costs()
        self.build_tariffs()

    def last_record(self):
        return self.data["history"][-1] if self.data["history"] else {}




    def build_main(self):

        self.day = tb.StringVar()
        self.night = tb.StringVar()
        self.water = tb.StringVar()

        hero = tb.Frame(self.tab_main)
        hero.pack(fill="x", pady=(25,10))

        tb.Label(
            hero,
            text="🏠 Комунальні платежі",
            font=("Bahnschrift SemiBold", 38)
        ).pack()

        tb.Label(
            hero,
            text="Ваш дім. Ваші витрати. Ваш контроль.",
            font=("Segoe UI", 14)
        ).pack(pady=(0,20))

        cards = tb.Frame(self.tab_main)
        cards.pack(pady=15)

        def card(parent, icon, title, var):
            f = tb.Frame(parent, bootstyle="dark")
            tb.Label(
                f,
                text=f"{icon}  {title}",
                font=("Bahnschrift SemiBold", 16)
            ).pack(anchor="w", padx=12, pady=(10,5))

            tb.Entry(
                f,
                textvariable=var,
                width=14,
                justify="center",
                font=("Bahnschrift SemiBold", 20)
            ).pack(padx=12, pady=(5,10))

            line = tb.Frame(f, height=2)
            line.pack(fill="x", padx=12, pady=5)

            tb.Label(
                f,
                text="Показник лічильника",
                font=("Segoe UI", 10)
            ).pack(anchor="w", padx=12, pady=(5,10))

            return f

        card(cards, "⚡", "День", self.day).grid(row=0,column=0,padx=18)
        card(cards, "🌙", "Ніч", self.night).grid(row=0,column=1,padx=18)
        card(cards, "🚰", "Вода", self.water).grid(row=0,column=2,padx=18)

        tariffs = tb.Frame(self.tab_main, bootstyle="dark")
        tariffs.pack(fill="x", padx=120, pady=20)

        tb.Label(
            tariffs,
            text="📋 Поточні тарифи",
            font=("Bahnschrift SemiBold", 14)
        ).pack(anchor="w", padx=15, pady=(10,5))

        tb.Label(
            tariffs,
            text=f"⚡ День: {TARIFFS['day']} грн     🌙 Ніч: {TARIFFS['night']} грн     🚰 Вода: {TARIFFS['water']} грн     🚿 Водовідв.: {TARIFFS['sewer']} грн",
            font=("Segoe UI", 12)
        ).pack(pady=(0,12))

        tb.Button(
            self.tab_main,
            text="💾 Розрахувати та зберегти",
            bootstyle="success",
            width=40,
            command=self.calculate
        ).pack(pady=20)

        last = self.last_record()

        status = tb.Frame(self.tab_main, bootstyle="dark")
        status.pack(fill="x", padx=120, pady=20)

        tb.Label(
            status,
            text="📈 Останні дані",
            font=("Bahnschrift SemiBold", 14)
        ).pack(anchor="w", padx=15, pady=(10,5))

        tb.Label(
            status,
            text=f"📅 Останній запис: {last.get('date','-')}      💰 Останнє нарахування: {last.get('total',0)} грн",
            font=("Segoe UI", 12)
        ).pack(pady=(0,12))

        self.total_label = tb.Label(
            self.tab_main,
            text="Загальна сума: 0.00 грн",
            font=("Bahnschrift SemiBold", 30)
        )
        self.total_label.pack(pady=20)

    def build_meters(self):

        style = ttk.Style()

        style.configure("Treeview", font=("Consolas", 12), rowheight=42)
        style.configure("Treeview.Heading", font=("Bahnschrift SemiBold", 15))

        header = tb.Frame(self.tab_meters)
        header.pack(fill="x", padx=20, pady=(20, 10))

        tb.Label(
            header,
            text="📊 Таблиця показників",
            font=("Bahnschrift SemiBold", 32)
        ).pack(anchor="w")

        search_frame = tb.Frame(self.tab_meters)
        search_frame.pack(fill="x", padx=25, pady=(0,10))

        container = tb.Labelframe(
            self.tab_meters,
            text=" Історія показників "
        )
        container.pack(fill="both", expand=True, padx=15, pady=5)

        tb.Label(search_frame,text="🔍 Пошук по даті:").pack(side="left")
        tb.Entry(search_frame,textvariable=self.search_meters,width=20).pack(side="left",padx=10)
        self.search_meters.trace_add("write", lambda *args: self.refresh_tables())

        cols = ("date", "day", "night", "water", "sewer")

        self.tree_meters = ttk.Treeview(
            container,
            columns=cols,
            show="headings"
        )

        headers = [
            "Дата",
            "День (кВт·год)",
            "Ніч (кВт·год)",
            "Вода (м³)",
            "Водовідв. (м³)"
        ]

        for c, h in zip(cols, headers):
            self.tree_meters.heading(c, text=h)

        self.tree_meters.column("date", width=130, anchor="center", stretch=True)
        self.tree_meters.column("day", width=180, anchor="center", stretch=True)
        self.tree_meters.column("night", width=180, anchor="center", stretch=True)
        self.tree_meters.column("water", width=130, anchor="center", stretch=True)
        self.tree_meters.column("sewer", width=130, anchor="center", stretch=True)
        
        scroll_y_m = ttk.Scrollbar(container, orient="vertical", command=self.tree_meters.yview)
        self.tree_meters.configure(yscrollcommand=scroll_y_m.set)

        scroll_y_m.pack(side="right", fill="y")
        self.tree_meters.pack(fill="both", expand=True, padx=15, pady=15)

        self.delta_label = tb.Label(
            container,
            text="⚡ День: - | 🌙 Ніч: - | 💧 Вода: - | 🚰 Водовідв.: -",
            font=("Bahnschrift", 12)
        )
        self.delta_label.pack(pady=(5, 10))

        self.tree_meters.bind(
            "<<TreeviewSelect>>",
            self.show_meter_difference
        )

        self.refresh_tables()


    def build_costs(self):

        style = ttk.Style()

        style.configure("Treeview", font=("Consolas", 12), rowheight=42)
        style.configure("Treeview.Heading", font=("Bahnschrift SemiBold", 15))

        header = tb.Frame(self.tab_costs)
        header.pack(fill="x", padx=20, pady=(20, 10))

        tb.Label(
            header,
            text="💰 Таблиця нарахувань",
            font=("Bahnschrift SemiBold", 32)
        ).pack(anchor="w")

        search_frame = tb.Frame(self.tab_costs)
        search_frame.pack(fill="x", padx=25, pady=(0,10))

        container = tb.Labelframe(
            self.tab_costs,
            text=" Історія нарахувань "
        )
        container.pack(fill="both", expand=True, padx=15, pady=5)

        tb.Label(search_frame,text="🔍 Пошук по даті:").pack(side="left")
        tb.Entry(search_frame,textvariable=self.search_costs,width=20).pack(side="left",padx=10)
        self.search_costs.trace_add("write", lambda *args: self.refresh_tables())


        cols = ("date","dayc","nightc","waterc","sewerc","mg","vat","dom","total")

        self.tree_costs = ttk.Treeview(
            container,
            columns=cols,
            show="headings",
            height=8
        )

        headers = [
            "Дата","День грн","Ніч грн","Вода грн",
            "Водовідв. грн","Управління","ПДВ","Домофон","Разом"
        ]

        for c,h in zip(cols,headers):
            self.tree_costs.heading(c,text=h)

        self.tree_costs.column("date", width=100, anchor="center")
        self.tree_costs.column("dayc", width=90, anchor="center")
        self.tree_costs.column("nightc", width=90, anchor="center")
        self.tree_costs.column("waterc", width=90, anchor="center")
        self.tree_costs.column("sewerc", width=100, anchor="center")
        self.tree_costs.column("mg", width=100, anchor="center")
        self.tree_costs.column("vat", width=70, anchor="center")
        self.tree_costs.column("dom", width=80, anchor="center")
        self.tree_costs.column("total", width=90, anchor="center")

        scroll_y_c = ttk.Scrollbar(container, orient="vertical", command=self.tree_costs.yview)

        self.tree_costs.configure(
            yscrollcommand=scroll_y_c.set
        )

        scroll_y_c.pack(side="right", fill="y")
        self.tree_costs.pack(fill="both", expand=True, padx=15, pady=15)

        bottom = tb.Frame(self.tab_costs)
        bottom.pack(fill="x", pady=(0, 8))

        tb.Button(
            bottom,
            text="📊 Експорт Excel",
            bootstyle="success",
            command=self.export_excel
        ).pack()

        stats = tb.Frame(self.tab_costs)
        stats.pack(fill="x", padx=20, pady=(0,8))

        card1 = tb.Labelframe(stats, text=" Нарахувань ")
        card1.pack(side="left", padx=15)

        card2 = tb.Labelframe(stats, text=" Фінанси ")
        card2.pack(side="left", padx=15)

        self.records_card = tb.Label(card1, text="0",
                                     font=("Bahnschrift SemiBold",18))
        self.records_card.pack(padx=30, pady=20)

        self.total_card = tb.Label(card2, text="0 грн",
                                   font=("Bahnschrift SemiBold",18))
        self.total_card.pack(padx=30, pady=20)

        self.refresh_tables()



    def build_tariffs(self):

        header = tb.Frame(self.tab_tariffs)
        header.pack(fill="x", padx=20, pady=(20,10))

        tb.Label(
            header,
            text="⚙️ Налаштування тарифів",
            font=("Bahnschrift SemiBold", 28)
        ).pack(anchor="w")

        self.tariff_entries = {}

        cards = tb.Frame(self.tab_tariffs)
        cards.pack(fill="x", padx=30, pady=20)

        sections = [
            ("⚡ Електроенергія", [("day","Світло день"),("night","Світло ніч")]),
            ("💧 Вода", [("water","Вода"),("sewer","Водовідведення")]),
            ("🏢 Додатково", [("management","Управління"),("vat","ПДВ"),("intercom","Домофон")]),
        ]

        for col, (title, items) in enumerate(sections):
            frame = tb.Labelframe(cards, text=f" {title} ")
            frame.grid(row=0, column=col, padx=15, sticky="nsew")

            for r, (key, label) in enumerate(items):
                tb.Label(frame, text=label).grid(row=r, column=0, padx=10, pady=8, sticky="w")

                e = tb.Entry(frame, width=12)
                e.insert(0, str(TARIFFS[key]))
                e.grid(row=r, column=1, padx=10, pady=8)

                self.tariff_entries[key] = e

        for i in range(3):
            cards.grid_columnconfigure(i, weight=1)

        tb.Button(
            self.tab_tariffs,
            text="💾 Зберегти зміни",
            bootstyle="success",
            width=25,
            command=self.save_tariffs
        ).pack(pady=20)

        tb.Label(
            self.tab_tariffs,
            text="✓ Тарифи зберігаються між запусками програми\n✓ Історія платежів не перераховується",
            font=("Segoe UI", 10)
        ).pack()
        
    def update_tariff_info(self):
        pass

    def save_tariffs(self):

        try:
            for key, entry in self.tariff_entries.items():
                TARIFFS[key] = float(entry.get())

            self.data["tariffs"] = TARIFFS.copy()
            save_data(self.data)

            messagebox.showinfo(
                "Готово",
                "Тарифи збережено. Після перезапуску програми вони залишаться без змін."
            )

        except ValueError:
            messagebox.showerror(
                "Помилка",
                "Усі тарифи повинні бути числовими значеннями."
            )

    def calculate(self):
        try:
            day = int(self.day.get())
            night = int(self.night.get())
            water = int(self.water.get())
        except:
            messagebox.showerror("Помилка", "Введіть числові значення")
            return

        if not self.data["history"]:
            self.data["history"].append({
                "date": datetime.now().strftime("%d.%m.%Y"),
                "day": day,
                "night": night,
                "water": water,
                "total": 0
            })
            save_data(self.data)
            self.refresh_tables()
            return

        prev = self.data["history"][-1]

        du = day - prev["day"]
        nu = night - prev["night"]
        wu = water - prev["water"]

        day_cost = du * TARIFFS["day"]
        night_cost = nu * TARIFFS["night"]
        water_cost = wu * TARIFFS["water"]
        sewer_cost = wu * TARIFFS["sewer"]

        total = (
            day_cost + night_cost + water_cost + sewer_cost +
            TARIFFS["management"] + TARIFFS["vat"] + TARIFFS["intercom"]
        )

        rec = {
            "date": datetime.now().strftime("%d.%m.%Y"),
            "day": day,
            "night": night,
            "water": water,
            "sewer_used": wu,
            "day_cost": round(day_cost,2),
            "night_cost": round(night_cost,2),
            "water_cost": round(water_cost,2),
            "sewer_cost": round(sewer_cost,2),
            "management": TARIFFS["management"],
            "vat": TARIFFS["vat"],
            "intercom": TARIFFS["intercom"],
            "total": round(total,2)
        }

        self.data["history"].append(rec)
        save_data(self.data)

        self.total_label.config(text=f"Загальна сума: {total:.2f} грн")
        self.refresh_tables()

    def refresh_tables(self):
        if hasattr(self, "tree_meters"):
            for i in self.tree_meters.get_children():
                self.tree_meters.delete(i)

            search = self.search_meters.get().strip() if hasattr(self,"search_meters") else ""
            for r in reversed(self.data["history"]):
                if search and search not in str(r.get("date","")):
                    continue
                self.tree_meters.insert("", "end", values=(
                    r.get("date",""),
                    r.get("day",""),
                    r.get("night",""),
                    r.get("water",""),
                    r.get("sewer_used",0)
                ))

            if hasattr(self, "records_card"):
                self.records_card.config(text=f"Всього записів: {len(self.data['history'])}")

            if hasattr(self, "total_card"):
                total_paid = sum(float(x.get("total", 0)) for x in self.data["history"])
                self.total_card.config(text=f"Всього сплачено: {total_paid:.2f} грн")

        if hasattr(self, "records_card"):
            self.records_card.config(
                text=f"Всього записів: {len(self.data['history'])}"
            )

        if hasattr(self, "total_card"):
            total_paid = sum(
                float(x.get("total", 0))
                for x in self.data["history"]
            )
            self.total_card.config(
                text=f"Всього сплачено: {total_paid:.2f} грн"
            )

        if hasattr(self, "tree_costs"):
            for i in self.tree_costs.get_children():
                self.tree_costs.delete(i)

            search = self.search_costs.get().strip() if hasattr(self,"search_costs") else ""
            for r in reversed(self.data["history"]):
                if search and search not in str(r.get("date","")):
                    continue
                self.tree_costs.insert("", "end", values=(
                    r.get("date",""),
                    r.get("day_cost",""),
                    r.get("night_cost",""),
                    r.get("water_cost",""),
                    r.get("sewer_cost",""),
                    r.get("management",""),
                    r.get("vat",""),
                    r.get("intercom",""),
                    r.get("total","")
                ))


    def show_meter_difference(self, event=None):

        selected = self.tree_meters.selection()

        if not selected:
            return

        values = self.tree_meters.item(selected[0])["values"]

        current_date = str(values[0])

        history = self.data["history"]

        idx = None
        for i, row in enumerate(history):
            if row.get("date") == current_date:
                idx = i
                break

        if idx is None or idx == 0:
            self.delta_label.config(
                text="Перший запис — немає попередніх даних"
            )
            return

        cur = history[idx]
        prev = history[idx - 1]

        day_diff = int(cur.get("day", 0)) - int(prev.get("day", 0))
        night_diff = int(cur.get("night", 0)) - int(prev.get("night", 0))
        water_diff = int(cur.get("water", 0)) - int(prev.get("water", 0))

        self.delta_label.config(
            text=(
                f"⚡ День: +{day_diff} кВт·год | "
                f"🌙 Ніч: +{night_diff} кВт·год | "
                f"💧 Вода: +{water_diff} м³ | "
                f"🚰 Водовідв.: +{water_diff} м³"
            )
        )

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.append(["Дата","День","Ніч","Вода","Разом"])

        for r in self.data["history"]:
            ws.append([
                r.get("date"),
                r.get("day"),
                r.get("night"),
                r.get("water"),
                r.get("total")
            ])

        wb.save(path)
        messagebox.showinfo("Готово", "Excel збережено")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    App().run()
