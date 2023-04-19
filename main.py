import tkinter as tk
from tkinter import ttk
from openpyxl import Workbook


class ConstructionEstimator(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Подсчет сметы")
        self.geometry("800x500")

        self.subject_var = tk.StringVar()
        self.quantity_var = tk.DoubleVar()
        self.value_var = tk.StringVar()
        self.price_var = tk.DoubleVar()

        self.create_widgets()

    def create_widgets(self):
        frame1 = ttk.Frame(self)
        frame1.pack(fill=tk.X, pady=5)

        frame3 = ttk.Frame(self)
        frame3.pack(fill=tk.X, pady=5)



        ttk.Label(frame1, text="Предмет:").pack(side=tk.LEFT, padx=5)
        subject_entry = ttk.Entry(frame1, textvariable=self.subject_var)
        subject_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(frame1, text="Количество:").pack(side=tk.LEFT, padx=5)
        quantity_entry = ttk.Entry(frame1, textvariable=self.quantity_var)
        quantity_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(frame1, text="Ед.Измерения:").pack(side=tk.LEFT, padx=5)
        value_entry = ttk.Entry(frame1, textvariable=self.value_var)
        value_entry.pack(side=tk.LEFT, padx=5)

        ttk.Label(frame1, text="Цена за еденицу:").pack(side=tk.LEFT, padx=5)
        price_entry = ttk.Entry(frame1, textvariable=self.price_var)
        price_entry.pack(side=tk.LEFT, padx=5)

        ttk.Button(frame3, text="Add", command=self.add_item).pack(side=tk.LEFT, padx=20)
        ttk.Button(frame3, text="Edit", command=self.edit_item).pack(side=tk.LEFT, padx=20)
        ttk.Button(frame3, text="Del", command=self.delete_item).pack(side=tk.LEFT, padx=20)
        ttk.Button(frame3, text="Export", command=self.export_to_excel).pack(side=tk.LEFT, padx=20)
        ttk.Button(frame3, text="Calculate", command=self.calculate_total).pack(side=tk.LEFT, padx=20)

        self.total_amount_label = ttk.Label(frame1)
        self.total_amount_label.pack(side=tk.RIGHT, padx=10)



        frame2 = ttk.Frame(self)
        frame2.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(frame2, columns=("subject", "quantity","value", "price"), show="headings")
        self.tree.heading("subject", text="Предмет")
        self.tree.heading("quantity", text="Кол-во")
        self.tree.heading("value", text="Ед.Измерения")
        self.tree.heading("price", text="Цена за еденицу")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def add_item(self):
        subject = self.subject_var.get()
        quantity = self.quantity_var.get()
        value = self.value_var.get()
        price = self.price_var.get()

        if subject and quantity > 0 and price > 0:
            self.tree.insert("", tk.END, values=(subject, quantity, value, price))

            # self.subject_var.set("")
            # self.quantity_var.set(0)
            # self.value_var.set("")
            # self.price_var.set(0)


    def edit_item(self):
        selected_item = self.tree.focus()

        if not selected_item:
            return

        subject = self.subject_var.get()
        quantity = self.quantity_var.get()
        value = self.value_var.get()
        price = self.price_var.get()

        if subject and quantity > 0 and price > 0:
            self.tree.item(selected_item, values=(subject, quantity, value, price))

    def delete_item(self):
        selected_item = self.tree.focus()
        if selected_item:
            self.tree.delete(selected_item)

    def export_to_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Estimation"
        ws.append(["Предмет", "Кол-во", "Ед.Измерения", "Цена за единицу"])

        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            ws.append(values)

        wb.save("Отчет.xlsx")

    def calculate_total(self):
        total_amount = 0

        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            total_amount += float(values[1]) * float(values[3])

        self.tree.insert("", tk.END, values=(f"Сумма: {total_amount}тг"))


if __name__ == "__main__":
    app = ConstructionEstimator()
    app.mainloop()
