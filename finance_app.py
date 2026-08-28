import sqlite3
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class FinanceDatabase:
    def __init__(self, db_name='finance.db'):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._create_tables()

    def _create_tables(self):
        """Создаем таблицы, если их нет"""
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                category TEXT NOT NULL,
                amount REAL NOT NULL,
                type TEXT NOT NULL,  -- 'income' или 'expense'
                description TEXT
            )
        ''')
        self.conn.commit()

    def add_transaction(self, category, amount, type_transaction, description=""):
        """Добавляет новую операцию (доход или расход)"""
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.cursor.execute('''
            INSERT INTO transactions (date, category, amount, type, description)
            VALUES (?, ?, ?, ?, ?)
        ''', (date, category, amount, type_transaction, description))
        self.conn.commit()

    def get_all_transactions(self):
        """Возвращает все операции"""
        self.cursor.execute('SELECT * FROM transactions ORDER BY date DESC')
        return self.cursor.fetchall()

    def delete_transaction(self, transaction_id):
        """Удаляет операцию по ID"""
        self.cursor.execute('DELETE FROM transactions WHERE id = ?', (transaction_id,))
        self.conn.commit()

    def get_balance(self):
        """Подсчет текущего баланса"""
        self.cursor.execute('''
            SELECT 
                SUM(CASE WHEN type = 'income' THEN amount ELSE 0 END) -
                SUM(CASE WHEN type = 'expense' THEN amount ELSE 0 END)
            FROM transactions
        ''')
        return self.cursor.fetchone()[0] or 0

    def get_expenses_by_category(self):
        """Получает расходы, сгруппированные по категориям (для графика)"""
        self.cursor.execute('''
            SELECT category, SUM(amount) as total
            FROM transactions
            WHERE type = 'expense'
            GROUP BY category
            ORDER BY total DESC
        ''')
        return self.cursor.fetchall()

    def export_to_excel(self, filename='finance_report.xlsx'):
        """Экспортирует все данные в Excel"""
        self.cursor.execute('SELECT * FROM transactions')
        rows = self.cursor.fetchall()
        
        # Преобразуем в DataFrame
        df = pd.DataFrame(rows, columns=['ID', 'Дата', 'Категория', 'Сумма', 'Тип', 'Описание'])
        
        # Меняем тип на русский для красоты
        df['Тип'] = df['Тип'].map({'income': 'Доход', 'expense': 'Расход'})
        
        df.to_excel(filename, index=False)
        return filename

    def close(self):
        self.conn.close()


class FinanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Учет финансов")
        self.root.geometry("800x600")
        
        self.db = FinanceDatabase()
        
        # --- Вкладки ---
        self.tab_control = ttk.Notebook(root)
        
        self.tab1 = ttk.Frame(self.tab_control)  # Вкладка операций
        self.tab2 = ttk.Frame(self.tab_control)  # Вкладка графиков
        
        self.tab_control.add(self.tab1, text='📝 Операции')
        self.tab_control.add(self.tab2, text='📊 Графики')
        self.tab_control.pack(expand=1, fill='both')
        
        # === Вкладка 1: Операции ===
        self._setup_tab1()
        self._setup_tab1_buttons()
        self.refresh_table()
        
        # === Вкладка 2: Графики ===
        self._setup_tab2()
        
    def _setup_tab1(self):
        """Настройка полей ввода"""
        frame = ttk.LabelFrame(self.tab1, text="Добавить операцию")
        frame.pack(pady=10, padx=10, fill='x')
        
        # Категория
        ttk.Label(frame, text="Категория:").grid(row=0, column=0, padx=5, pady=5)
        self.category_var = tk.StringVar()
        self.category_entry = ttk.Entry(frame, textvariable=self.category_var)
        self.category_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Сумма
        ttk.Label(frame, text="Сумма (₽):").grid(row=0, column=2, padx=5, pady=5)
        self.amount_var = tk.StringVar()
        self.amount_entry = ttk.Entry(frame, textvariable=self.amount_var)
        self.amount_entry.grid(row=0, column=3, padx=5, pady=5)
        
        # Тип
        ttk.Label(frame, text="Тип:").grid(row=0, column=4, padx=5, pady=5)
        self.type_var = tk.StringVar(value="expense")
        type_combo = ttk.Combobox(frame, textvariable=self.type_var, state="readonly")
        type_combo['values'] = ("expense", "income")
        type_combo.grid(row=0, column=5, padx=5, pady=5)
        
        # Описание
        ttk.Label(frame, text="Описание:").grid(row=1, column=0, padx=5, pady=5)
        self.desc_var = tk.StringVar()
        self.desc_entry = ttk.Entry(frame, textvariable=self.desc_var)
        self.desc_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky='we')
        
    def _setup_tab1_buttons(self):
        """Кнопки"""
        frame = ttk.Frame(self.tab1)
        frame.pack(pady=5)
        
        ttk.Button(frame, text="Добавить", command=self.add_transaction).pack(side='left', padx=5)
        ttk.Button(frame, text="Удалить выбранное", command=self.delete_transaction).pack(side='left', padx=5)
        ttk.Button(frame, text="Экспорт в Excel", command=self.export_excel).pack(side='left', padx=5)
        ttk.Button(frame, text="Обновить", command=self.refresh_table).pack(side='left', padx=5)
        
        # Таблица
        self.tree = ttk.Treeview(self.tab1, columns=('ID', 'Дата', 'Категория', 'Сумма', 'Тип', 'Описание'), show='headings')
        self.tree.heading('ID', text='ID')
        self.tree.heading('Дата', text='Дата')
        self.tree.heading('Категория', text='Категория')
        self.tree.heading('Сумма', text='Сумма')
        self.tree.heading('Тип', text='Тип')
        self.tree.heading('Описание', text='Описание')
        
        self.tree.column('ID', width=40)
        self.tree.column('Сумма', width=80)
        self.tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Баланс
        self.balance_label = ttk.Label(self.tab1, text="Баланс: 0 ₽", font=("Arial", 14, "bold"))
        self.balance_label.pack(pady=10)
        
    def _setup_tab2(self):
        """Настройка графика"""
        self.figure = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.tab2)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
        
        ttk.Button(self.tab2, text="Показать расходы по категориям", command=self.draw_chart).pack(pady=5)
        
    # --- Логика ---
    def add_transaction(self):
        try:
            category = self.category_var.get().strip()
            amount = float(self.amount_var.get())
            type_trans = self.type_var.get()
            desc = self.desc_var.get().strip()
            
            if not category or amount <= 0:
                raise ValueError("Введите категорию и сумму больше 0")
            
            self.db.add_transaction(category, amount, type_trans, desc)
            
            # Очищаем поля
            self.category_var.set("")
            self.amount_var.set("")
            self.desc_var.set("")
            
            self.refresh_table()
            messagebox.showinfo("Успех", "Операция добавлена!")
        except ValueError as e:
            messagebox.showerror("Ошибка", str(e))
        
    def delete_transaction(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Внимание", "Выберите операцию для удаления")
            return
        
        item = self.tree.item(selected[0])
        transaction_id = item['values'][0]
        
        self.db.delete_transaction(transaction_id)
        self.refresh_table()
        messagebox.showinfo("Успех", "Операция удалена!")
        
    def refresh_table(self):
        """Обновляет таблицу и баланс"""
        # Очищаем таблицу
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Загружаем данные
        for row in self.db.get_all_transactions():
            self.tree.insert('', 'end', values=row)
        
        # Обновляем баланс
        balance = self.db.get_balance()
        self.balance_label.config(text=f"Баланс: {balance:.2f} ₽")
        
    def export_excel(self):
        filename = self.db.export_to_excel()
        messagebox.showinfo("Успех", f"Файл сохранен: {filename}")
        
    def draw_chart(self):
        """Рисует круговую диаграмму расходов"""
        data = self.db.get_expenses_by_category()
        
        if not data:
            messagebox.showinfo("Инфо", "Нет расходов для отображения")
            return
        
        # Очищаем график
        self.ax.clear()
        
        # Разделяем данные
        categories = [item[0] for item in data]
        amounts = [item[1] for item in data]
        
        # Рисуем
        self.ax.pie(amounts, labels=categories, autopct='%1.1f%%')
        self.ax.set_title("Расходы по категориям")
        
        # Обновляем canvas
        self.canvas.draw()


if __name__ == "__main__":
    root = tk.Tk()
    app = FinanceApp(root)
    root.mainloop()