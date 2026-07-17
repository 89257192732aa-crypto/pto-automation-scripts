import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def run_calculator():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    clear_screen()
    print('*'*50)
    print("          SMETA CONVERTER v2.2 (TEMPLATE)")
    print('*'*50)

    try:
        print("\n[1] Выберите файл сметы (откуда берем данные)...")
        smeta_path = filedialog.askopenfilename(title="Открыть смету", filetypes=[("Excel", "*.xlsx *.xls")])

        if not smeta_path:
            print("Выбор отменен.")
            return

        df_smeta = pd.read_excel(smeta_path, header=None)
        total_rows = len(df_smeta)
        
        print(f"Файл загружен. Всего строк: {total_rows}")

        print("\n[2] Введите параметры:")
        prefix_in = input("Префикс (Enter = Акт скрытых работ): ")
        prefix = prefix_in if prefix_in else "Акт освидетельствования скрытых работ. "
        
        while True:
            col_letter = input("Буква колонки (A, B, C...): ").upper() or "B"
            try:
                col_idx = openpyxl.utils.column_index_from_string(col_letter) - 1
                break
            except:
                print("Ошибка! Введите букву латиницей.")

        while True:
            try:
                row_start_in = input("С какой строки начать? (Enter = 1): ")
                row_start = int(row_start_in) if row_start_in else 1
                
                row_end_in = input(f"На какой закончить? (Enter = {total_rows}): ")
                row_end = int(row_end_in) if row_end_in else total_rows
                break
            except:
                print("Ошибка! Введите число.")

        # --- ОБРАБОТКА ---
        final_works = []
        for i in range(row_start - 1, row_end):
            if i < total_rows:
                val = df_smeta.iloc[i, col_idx]
                if pd.notna(val) and str(val).strip() != "":
                    final_works.append(f"{prefix}{val}")

        # --- СОХРАНЕНИЕ В ШАБЛОН ---
        if not final_works:
            print("\nДанные не найдены.")
        else:
            print(f"\nСобрано строк: {len(final_works)}")
            print("\n[3] Выберите файл ШАБЛОНА (который нужно заполнить)...")
            root.update()
            template_path = filedialog.askopenfilename(title="Открыть шаблон", filetypes=[("Excel", "*.xlsx")])

            if template_path:
                # Загружаем шаблон с сохранением стилей
                wb = openpyxl.load_workbook(template_path)
                sheet = wb.active 

                # НАСТРОЙКИ ВСТАВКИ (можно поменять на нужные)
                start_row_sh = 10 
                start_col_sh = 2

                for index, work_text in enumerate(final_works):
                    sheet.cell(row=start_row_sh + index, column=start_col_sh).value = work_text
                
                print("[4] Сохранение результата...")
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile="Ready_Smeta.xlsx",
                    filetypes=[("Excel", "*.xlsx")]
                )

                if output_path:
                    wb.save(output_path)
                    print("\n" + "*"*50)
                    print("ГОТОВО! Шаблон заполнен и сохранен.")
                    print(f"Путь: {output_path}")
                    print("*"*50)
                    os.startfile(os.path.dirname(output_path))
                else:
                    print("Сохранение отменено.")
            else:
                print("Шаблон не выбран.")

    except Exception as e:
        print(f"\nОШИБКА: {e}")
    finally:
        root.destroy()

if __name__ == "__main__":
    run_calculator()
    input("\nНажмите ENTER для выхода...")
