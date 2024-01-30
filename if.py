import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pyodbc as pd
import numpy as np
import tkinter.simpledialog

# Устанавливаем строку подключения к базе данных
connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/user/Desktop/n/rs.mdb'

# Функция для извлечения данных иерархии по ID экспертизы
def fetch_hierarchy_data(expertise_id):
    conn = pd.connect(connection_string)
    cursor = conn.cursor()
    cursor.execute('SELECT id_node, id_parent, id_e FROM Hierarchy WHERE id_e = ?', expertise_id)
    hierarchy_data = cursor.fetchall()
    cursor.close()
    conn.close()
    return hierarchy_data

# Функция для извлечения матрицы сравнений по ID узла из базы данных
def fetch_matrix_data(node_id):
    conn = pd.connect(connection_string)
    cursor = conn.cursor()
    cursor.execute('SELECT pos, value FROM Matrix WHERE id_n = ?', node_id)
    matrix_data = cursor.fetchall()
    cursor.close()
    conn.close()
    return matrix_data

# Функция для извлечения данных экспертизы
def fetch_expertise_names():
    conn = pd.connect(connection_string)
    cursor = conn.cursor()
    cursor.execute('SELECT id, name FROM Elements')
    expertise_data = cursor.fetchall()
    cursor.close()
    conn.close()
    return expertise_data

# Функция для создания матрицы парных сравнений на основе значений
def create_comparison_matrix(values):
    n = int(len(values) ** 0.5)  # Определяем размер матрицы (предполагая, что она квадратная)
    matrix = [[0] * n for _ in range(n)]

    for i in range(n):
        for j in range(n):
            if i == j:
                matrix[i][j] = 1  # Заполняем диагональные элементы значением 1
            else:
                matrix[i][j] = values[i * n + j]

    return matrix

# Функция для отображения матрицы парных сравнений по ID узла
def show_comparison_matrix(node_id):
    matrix_data = fetch_matrix_data(node_id)
    if matrix_data:
        print(f"Matrix data for node {node_id}: {matrix_data}")  # Отладочный вывод
    if matrix_data:
        n = int(len(matrix_data) ** 0.5)  # Определяем размер матрицы

        matrix_values = [value for _, value in matrix_data]
        comparison_matrix = create_comparison_matrix(matrix_values)

        # Создаем новое окно для отображения матрицы парных сравнений
        matrix_window = tk.Toplevel()
        matrix_window.title(f"Матрица парных сравнений (ID узла: {node_id})")

        # Создаем рамку в окне для отображения элементов
        frame = ttk.Frame(matrix_window)
        frame.pack(fill="both", expand=True)

        # Создаем Treeview для отображения матрицы
        matrix_tree = ttk.Treeview(frame)
        matrix_tree["columns"] = [str(i + 1) for i in range(n)]

        for i in range(n):
            matrix_tree.column(str(i + 1), width=100, anchor="center")
            matrix_tree.heading(str(i + 1), text=str(i + 1))

        for i in range(n):
            values = comparison_matrix[i]
            matrix_tree.insert("", "end", text=str(i + 1), values=values)

        matrix_tree.pack(fill="both", expand=True)

        # Вычисление и вывод собственного вектора, отношения компонентов и других параметров
        eigenvector = calculate_eigenvector(comparison_matrix)
        eigenvalues, consistency_index, random_index, consistency_ratio = calculate_consistency(eigenvector, comparison_matrix)
        # Создаем кнопку "Ок"
        ok_button = tk.Button(frame, text="Ок", command=on_ok_button_clicked)
        ok_button.pack()
        # Вывод собственного вектора
        eigenvector_label = tk.Label(frame, text="Собственный вектор")
        eigenvector_label.pack()

        eigenvector_tree = ttk.Treeview(frame, columns=["Компонента", "Значение"], show="headings")
        eigenvector_tree.pack(fill="both", expand=True)

        eigenvector_tree.heading("Компонента", text="Компонента")
        eigenvector_tree.heading("Значение", text="Значение")

        for i, value in enumerate(eigenvector):
            eigenvector_tree.insert("", "end", values=(i + 1, value))

        # Вывод разницы
        difference_label = tk.Label(frame, text="Разница")
        difference_label.pack()

        difference_tree = ttk.Treeview(frame, columns=[str(i + 1) for i in range(n)], show="headings")
        difference_tree.pack(fill="both", expand=True)

        difference_tree.heading("#0", text="Компонента")

        for i in range(n):
            difference_tree.heading(str(i + 1), text=str(i + 1))
            difference_tree.column(str(i + 1), width=100, stretch=tk.NO)

        for i in range(n):
            difference_tree.insert("", "end", values=(i + 1, *calculate_difference(comparison_matrix, eigenvector)[i]))

        # Вывод собственных чисел и параметров оценки согласованности
        eigenvalues_label = tk.Label(frame, text="Собственные числа:")
        eigenvalues_label.pack()
        eigenvalues_label = tk.Label(frame, text=eigenvalues)
        eigenvalues_label.pack()

        ci_label = tk.Label(frame, text="Индекс согласованности:")
        ci_label.pack()
        ci_label = tk.Label(frame, text=consistency_index)
        ci_label.pack()

        ri_label = tk.Label(frame, text="Случайный индекс:")
        ri_label.pack()
        ri_label = tk.Label(frame, text=random_index)
        ri_label.pack()

        cr_label = tk.Label(frame, text="Отношение согласованности:")
        cr_label.pack()
        cr_label = tk.Label(frame, text=consistency_ratio)
        cr_label.pack()

    else:
        messagebox.showerror("Ошибка", "Нет данных для выбранного узла.")
def on_ok_button_clicked():
    # Получаем новые значения для матрицы
    new_values = get_new_matrix_values()

    # Обновляем все значения на основе новой матрицы
    update_values_based_on_matrix(new_values)

def get_new_matrix_values():
    # Создаем диалоговое окно для ввода новых значений матрицы
    new_values = tkinter.simpledialog.askstring("Изменение значений", "Введите новые значения матрицы (разделите числа запятыми):")

    # Преобразуем строку в список чисел
    new_values = [float(x.strip()) for x in new_values.split(",")]

    return new_values

def update_values_based_on_matrix(new_values):
    # Обновляем отношение компонент собственного вектора и другие значения на основе новых значений матрицы
    comparison_matrix = create_comparison_matrix(new_values)

    eigenvector = calculate_eigenvector(comparison_matrix)
    _, consistency_index, random_index, consistency_ratio = calculate_consistency(eigenvector, comparison_matrix)

    update_displayed_values(eigenvector, consistency_index, random_index, consistency_ratio)

def update_displayed_values(eigenvector, consistency_index, random_index, consistency_ratio):
    # Обновляем отображаемые значения в интерфейсе
    eigenvector_tree.delete(*eigenvector_tree.get_children())
    for i, value in enumerate(eigenvector):
        eigenvector_tree.insert("", "end", values=(i + 1, value))

    difference_tree.delete(*difference_tree.get_children())
    for i in range(len(eigenvector)):
        difference_tree.insert("", "end", values=(i + 1, *calculate_difference(comparison_matrix, eigenvector)[i]))

    eigenvalues_label.config(text="Собственные числа: " + str(np.max(np.linalg.eigvals(comparison_matrix))))
    ci_label.config(text="Индекс согласованности: " + str(consistency_index))
    ri_label.config(text="Случайный индекс: " + str(random_index))
    cr_label.config(text="Отношение согласованности: " + str(consistency_ratio))

# Функция для обновления дерева с экспертизами
def populate_tree(tree, expertise_id):
    # Очистим дерево перед обновлением
    tree.delete(*tree.get_children())
    hierarchy_data = fetch_hierarchy_data(expertise_id)

    # Создадим словарь узлов для построения дерева
    nodes = {id_node: {'parent_id': id_parent} for id_node, id_parent, _ in hierarchy_data}

    for id_node, node in nodes.items():
        parent_id = node['parent_id']
        if parent_id not in nodes:
            tree.insert("", "end", text=f"{id_node}", values=(parent_id, "Parent"))
        else:
            parent_node = nodes[parent_id]
            if "Parent" in tree.item(str(parent_id), option="values"):
                tree.insert(str(parent_id), "end", text=f"{id_node}", values=(parent_id, "Child"))
            else:
                tree.insert(str(parent_id), "end", text=f"{id_node}")

    def on_select(event):
        selected_item = tree.selection()[0]
        node_id = int(tree.item(selected_item, option="text"))
        show_matrix_by_id_n(node_id)

    # При выборе узла в дереве отобразим матрицу сравнений по ID узла
    tree.bind("<<TreeviewSelect>>", on_select)


def extract_node_id(selected_node):
    match = re.search(r'\d+', selected_node)  # Находим все последовательности цифр в строке
    if match:
        node_id = match.group()  # Получаем найденное числовое значение
        return int(node_id)
    return None  # Если не найдено, возвращаем None или какой-то другой признак ошибки


def extract_comparison_values_by_id_n(id_n):
    values = []

    # Подключение к базе данных
    conn = pd.connect(connection_string)
    cursor = conn.cursor()

    # Извлечение данных из таблицы 'Matrix' для конкретного id_n
    cursor.execute('SELECT value FROM Matrix WHERE id_n = ?', id_n)
    values_data = cursor.fetchall()

    # Получаем значения "Value" и добавляем их в список
    values = [row[0] for row in values_data]

    # Закрытие соединения
    cursor.close()
    conn.close()

    return values

def calculate_eigenvector(matrix):
    # Вычисление собственного вектора
    eigenvalues, eigenvectors = np.linalg.eig(matrix)

    # Находим индекс максимального собственного значения
    max_eigenvalue_index = np.argmax(eigenvalues)

    # Извлекаем соответствующий собственный вектор
    eigenvector = eigenvectors[:, max_eigenvalue_index]

    # Нормализуем собственный вектор
    normalized_eigenvector = eigenvector / np.sum(eigenvector)

    return normalized_eigenvector

# Функция для вычисления параметров оценки согласованности
def calculate_consistency(eigenvector, matrix):
    n = len(matrix)

    # Собственные числа
    eigenvalues, _ = np.linalg.eig(matrix)

    # Максимальное собственное число
    max_eigenvalue = np.max(eigenvalues)

    # Индекс согласованности
    consistency_index = (max_eigenvalue - n) / (n - 1)

    # Случайный индекс
    random_index = 1.98 * (n - 2) / n

    # Отношение согласованности
    consistency_ratio = consistency_index / random_index

    return eigenvalues, consistency_index, random_index, consistency_ratio


def calculate_difference(matrix, eigenvector):
    n = len(matrix)
    difference_matrix = np.zeros((n, n))

    for i in range(n):
        for j in range(n):
            difference_matrix[i][j] = matrix[i][j] - np.real(eigenvector[i] / eigenvector[j])

    return difference_matrix
def your_improvement_function(current_values):
    n = int(len(current_values) ** 0.5)

    # Преобразуем текущие значения в матрицу
    current_matrix = np.array(current_values).reshape((n, n))

    # Рассчитываем собственный вектор и собственное число
    eigenvector = calculate_eigenvector(current_matrix)
    eigenvalue = np.max(np.linalg.eigvals(current_matrix))

    # Находим максимальное значение в собственном векторе
    max_value = np.max(eigenvector)

    # Заменяем минимальное значение в матрице собственным числом
    improved_matrix = current_matrix + (eigenvalue - max_value) * np.identity(n)

    # Возвращаем улучшенные значения
    return improved_matrix.flatten()

def show_matrix_by_id_n(id_n):
    # Извлекаем значения "Value" из базы данных для данного id_n
    values = extract_comparison_values_by_id_n(id_n)

    if values:
        n = int(len(values) ** 0.5)  # Определяем размер матрицы (предполагая, что она квадратная)

        # Создаем матрицу парных сравнений
        comparison_matrix = create_comparison_matrix(values)

        matrix_window = tk.Toplevel()
        matrix_window.title(f"Матрица парных сравнений (id_n = {id_n})")

        # Добавим полосы прокрутки
        matrix_canvas = tk.Canvas(matrix_window)
        matrix_scroll_y = tk.Scrollbar(matrix_window, orient="vertical", command=matrix_canvas.yview)
        matrix_scroll_y.pack(side="right", fill="y")
        matrix_canvas.pack(side="left", fill="both", expand=True)
        matrix_canvas.configure(yscrollcommand=matrix_scroll_y.set)

        frame = ttk.Frame(matrix_canvas)
        matrix_canvas.create_window((0, 0), window=frame, anchor="nw")

        # Создаем список заголовков столбцов для матрицы
        column_headings = [str(i + 1) for i in range(n)]
        matrix_tree = ttk.Treeview(frame, columns=column_headings, show="headings")
        matrix_tree.pack(fill="both", expand=True)

        for heading in column_headings:
            matrix_tree.heading(heading, text=heading)
            matrix_tree.column(heading, width=100, stretch=tk.NO)
        for i in range(n):
            matrix_tree.insert("", "end", values=values[i * n:(i + 1) * n])

        # Вычисление и вывод собственного вектора, отношения компонентов и других параметров
        eigenvalue = np.max(np.linalg.eigvals(comparison_matrix))
        eigenvector = calculate_eigenvector(comparison_matrix)
        _, consistency_index, random_index, consistency_ratio = calculate_consistency(eigenvector, comparison_matrix)

        # Вывод собственного вектора
        eigenvector_label = tk.Label(frame, text="Собственный вектор")
        eigenvector_label.pack()

        eigenvector_tree = ttk.Treeview(frame, columns=["Компонента", "Значение"], show="headings")
        eigenvector_tree.pack(fill="both", expand=True)

        eigenvector_tree.heading("Компонента", text="Компонента")
        eigenvector_tree.heading("Значение", text="Значение")

        for i, value in enumerate(eigenvector):
            eigenvector_tree.insert("", "end", values=(i + 1, value))

        # Вывод отношения компонентов собственного вектора в виде матрицы парных сравнений
        relation_matrix = np.zeros((n, n))
        for i in range(n):
            for j in range(n):
                relation_matrix[i][j] = np.real(eigenvector[i] / eigenvector[j])

        relation_matrix_label = tk.Label(frame, text="Матрица отношения компонентов собственного вектора")
        relation_matrix_label.pack()

        relation_matrix_tree = ttk.Treeview(frame, columns=[str(i + 1) for i in range(n)], show="headings")
        relation_matrix_tree.pack(fill="both", expand=True)

        for i in range(n):
            relation_matrix_tree.heading(str(i + 1), text=str(i + 1))
            relation_matrix_tree.column(str(i + 1), width=100, stretch=tk.NO)

        for i in range(n):
            values = ["{:.4f}".format(relation_matrix[i][j]) for j in range(n)]
            relation_matrix_tree.insert("", "end", values=values)

        # Вывод разницы между оценкой и отношением компонент собственного вектора
        difference_matrix = np.subtract(comparison_matrix, relation_matrix)

        difference_label = tk.Label(frame, text="Разница между оценкой и отношением компонент собственного вектора")
        difference_label.pack()

        difference_tree = ttk.Treeview(frame, columns=[str(i + 1) for i in range(n)], show="headings")
        difference_tree.pack(fill="both", expand=True)

        difference_tree.heading("#0", text="Компонента")

        for i in range(n):
            difference_tree.heading(str(i + 1), text=str(i + 1))
            difference_tree.column(str(i + 1), width=100, stretch=tk.NO)

        for i in range(n):
            values = ["{:.4f}".format(difference_matrix[i][j]) for j in range(n)]
            difference_tree.insert("", "end", values=values)

        # Вывод собственного числа и параметров оценки согласованности
        eigenvalue_label = tk.Label(frame, text="Собственное число:")
        eigenvalue_label.pack()

        eigenvalue_label = tk.Label(frame, text=eigenvalue)
        eigenvalue_label.pack()

        ci_label = tk.Label(frame, text="Индекс согласованности:")
        ci_label.pack()
        ci_label = tk.Label(frame, text=consistency_index)
        ci_label.pack()

        ri_label = tk.Label(frame, text="Случайный индекс:")
        ri_label.pack()
        ri_label = tk.Label(frame, text=random_index)
        ri_label.pack()

        cr_label = tk.Label(frame, text="Отношение согласованности:")
        cr_label.pack()
        cr_label = tk.Label(frame, text=consistency_ratio)
        cr_label.pack()

        # Добавляем кнопку для улучшения согласованности
        improve_button = tk.Button(frame, text="Улучшить согласованность", command=lambda: improve_consistency(id_n, matrix_tree))
        improve_button.pack()

        # Устанавливаем область прокрутки
        frame.update_idletasks()
        matrix_canvas.config(scrollregion=matrix_canvas.bbox("all"))

    else:
        messagebox.showerror("Ошибка", "Неверное значение id_n или матрица не существует.")

def improve_consistency(id_n, matrix_tree):
    # Получаем текущие значения матрицы из Treeview
    current_values = []
    for child_id in matrix_tree.get_children():
        values = matrix_tree.item(child_id, option="values")
        current_values.extend(values)

    # Преобразуем строки в числа
    current_values = [float(value) for value in current_values]

    # Проводим процедуру улучшения согласованности (заменяем значения и обновляем окно)
    improved_values = your_improvement_function(current_values)

    # Обновляем Treeview с новыми значениями
    for child_id, value in zip(matrix_tree.get_children(), improved_values):
        matrix_tree.item(child_id, values=value)

# Функция для обработки выбора экспертизы из выпадающего списка
def on_expertise_selected(event):
    selected_expertise_id = expertise_names[expertise_combobox.current()][0]
    new_window = tk.Toplevel(root)
    new_window.title("Дерево принятия решений")

    # Создаем Treeview для отображения иерархии
    tree = ttk.Treeview(new_window)
    tree["columns"] = ("Parent_ID")
    tree.column("#0", width=100)
    tree.heading("#0", text="ID")
    tree.column("Parent_ID", width=100)
    tree.heading("Parent_ID", text="Parent ID")
    tree.grid(row=1, column=0, pady=10, padx=10)

    # После выбора экспертизы из списка, заполним дерево с ее иерархией
    populate_tree(tree, selected_expertise_id)

    def on_select(event):
        selected_item = tree.selection()[0]
        values = tree.item(selected_item, option="values")
        show_matrix_by_id_n(int(values[0]))

    # При выборе узла в дереве отобразим матрицу сравнений по ID узла
    tree.bind("<<TreeviewSelect>>", on_select)


# Создаем основное окно
root = tk.Tk()
root.title("Просмотр иерархии")

# Получаем данные экспертизы
expertise_names = fetch_expertise_names()

# Создаем выпадающий список с экспертизами
expertise_combobox = ttk.Combobox(root, values=[name for _, name in expertise_names], state="readonly")
expertise_combobox.grid(row=0, column=0, pady=10, padx=10)
expertise_combobox.bind("<<ComboboxSelected>>", on_expertise_selected)

root.mainloop()