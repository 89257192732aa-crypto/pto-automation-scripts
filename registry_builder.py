import os
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Функция для красивой очистки экрана
def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def run_calculator():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    clear_screen()
    print("="*50)
    print("       🚀 SMETA CONVERTER v1.0 🚀")
    print("="*50)

    try:
        print("\n[ШАГ 1] Выберите файл сметы...")
        smeta_path = filedialog.askopenfilename(title="Выберите файл сметы", filetypes=[("Excel files", "*.xlsx *.xls")])

        if not smeta_path:
            print("❌ Файл не выбран!")
            return

        df_smeta = pd.read_excel(smeta_path, header=None)
        total_rows = len(df_smeta)
        
        clear_screen()
        print(f"📁 Файл: {os.path.basename(smeta_path)}")
        print(f"📊 Найдено строк: {total_rows}")
        print("-"*30)

        # 2. НАСТРОЙКИ
        prefix = input("\n📝 Введите текст ПРЕФИКСА (или Enter): ") or "Акт освидетельствования скрытых работ. "
        
        print("\n[ШАГ 2] Настройка колонок и строк")
        while True:
            col_letter = input("   👉 Буква колонки (A, B, C...): ").upper() or "B"
            try:
                col_idx = openpyxl.utils.column_index_from_string(col_letter) - 1
                break
            except: print("   ⚠️ Ошибка! Используйте латиницу.")

        while True:
            try:
                row_start = int(input("   👉 С какой строки начать? (Enter=7): ") or 7)
                row_end = int(input(f"   👉 На какой закончить? (Enter={total_rows}): ") or total_rows)
                break
            except: print("   ⚠️ Введите число цифрами.")

        # 3. СКЛЕИВАНИЕ
        final_works = []
        for i in range(row_start - 1, row_end):
            if i < total_rows:
                val = df_smeta.iloc[i, col_idx]
                if pd.notna(val) and str(val).strip() != "":
                    final_works.append(f"{prefix}{val}")

        # 4. СОХРАНЕНИЕ
        if not final_works:
            print("\nПусто! Нечего сохранять.")
        else:
            print(f"\n[ШАГ 3] Сохранение {len(final_works)} строк...")
            root.update()
            output_path = filedialog.asksaveasfilename(
                title="Куда сохранить результат?",
                defaultextension=".xlsx",
                initialfile="Готовый_реестр.xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if output_path:
                wb = openpyxl.Workbook()
                sheet = wb.active
                for index, work_text in enumerate(final_works):
                    sheet.cell(row=1 + index, column=1).value = work_text
                
                wb.save(output_path)
                print("\n" + "="*50)
                print(f"✅ УСПЕШНО СОХРАНЕНО!")
                print(f"📍 Путь: {output_path}")
                print("="*50)
                os.startfile(os.path.dirname(output_path))
            else:
                print("❌ Сохранение отменено.")

    except Exception as e:
        print(f"\n🆘 Ошибка: {e}")
    finally:
        root.destroy()

if __name__ == "__main__":
    run_calculator()
    print("\nРабота завершена.")
    input("Нажмите ENTER для выхода...")
