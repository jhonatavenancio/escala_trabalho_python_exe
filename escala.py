import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import random
from datetime import datetime, timedelta
import os


# Função para gerar uma lista de datas em um mês especifico
def generate_workdays(year, month):
    start_date = datetime(year, month, 1)
    end_date = (start_date + timedelta(days=31)).replace(day=1) - timedelta(days=1)

    workdays = [start_date + timedelta(days=x) for x in range((end_date - start_date).days + 1)]
    workdays = [day for day in workdays if day.weekday() < 5]

    return workdays

# Gerar a escala de trabalho para um mês com home office
def generate_schedule(employee_names, year, month):
    workdays = generate_workdays(year, month)

    df = pd.DataFrame(index=workdays, columns=employee_names)
    df.index.name = 'Dia'
    df.reset_index(inplace=True)  
    df['Dia'] = df['Dia'].dt.strftime("%d/%m/%Y") 

    # Agrupa dias úteis por semana e define 3 dias de home office por funcionário
    weeks = {}
    current_week = []
    last_weekday = -1
    week_number = 0

    for workday in workdays:
        weekday = workday.weekday()
        if weekday <= last_weekday:

            weeks[week_number] = current_week
            current_week = []
            week_number += 1
        current_week.append(workday)
        last_weekday = weekday
    weeks[week_number] = current_week

    # Define a escala de home office
    for employee in employee_names:
        for week in weeks.values():
            if len(week) >= 5:
                home_office_days = random.sample(week, 3)
            else:
                home_office_days = week[:3]

            # Marca os dias como 'Presencial' por padrão
            for day in week:
                df.loc[df['Dia'] == day.strftime("%d/%m/%Y"), employee] = "Presencial"

            # Define 3 dias de home office
            for day in home_office_days:
                df.loc[df['Dia'] == day.strftime("%d/%m/%Y"), employee] = "Home"

    return df


# Função para salvar a escala em arquivo Excel
def save_schedule(schedule):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    target_folder = os.path.expanduser("~/Documentos/home_presencial")
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    file_path = os.path.join(target_folder, f"escala_trabalho_{timestamp}.xlsx")
    schedule.to_excel(file_path, index=False)
    return file_path


# Função para gerar a escala e exibir uma mensagem de sucesso ou erro
def generate_and_save_schedule(employee_names_str, year, month):
    employee_names = [name.strip() for name in employee_names_str.split(',')]
    if len(employee_names) > 20:
        raise ValueError("Por favor, insira no máximo 20 nomes.")
    
    schedule = generate_schedule(employee_names, year, month)
    return save_schedule(schedule)


# Interface gráfica usando tkinter
class ScheduleApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Escala de Trabalho")
        
        self.label_employee_names = tk.Label(self, text="Nomes dos funcionários (separados por vírgula):")
        self.label_employee_names.grid(row=0, column=0, padx=10, pady=5)

        self.entry_employee_names = tk.Entry(self, width=50)
        self.entry_employee_names.grid(row=0, column=1, padx=10, pady=5)

        self.label_month = tk.Label(self, text="Mês:  (Abril: 4)")
        self.label_month.grid(row=1, column=0, padx=10, pady=5)

        self.entry_month = tk.Entry(self, width=11)
        self.entry_month.grid(row=1, column=1, padx=10, pady=5)

        self.label_year = tk.Label(self, text="Ano:  (2024)")
        self.label_year.grid(row=2, column=0, padx=10, pady=5)

        self.entry_year = tk.Entry(self, width=11)
        self.entry_year.grid(row=2, column=1, padx=10, pady=5)

        self.button_generate = tk.Button(self, text="Gerar Escala", command=self.generate_schedule)
        self.button_generate.grid(row=3, column=0, columnspan=2, pady=10)

    def generate_schedule(self):
        try:
            employee_names_str = self.entry_employee_names.get()
            month = int(self.entry_month.get())
            year = int(self.entry_year.get())
            file_path = generate_and_save_schedule(employee_names_str, year, month)
            messagebox.showinfo("Sucesso", f"Escala de trabalho salva em '{file_path}'")
        except Exception as e:
            messagebox.showerror("Erro", str(e))


if __name__ == "__main__":
    app = ScheduleApp()
    app.mainloop()
