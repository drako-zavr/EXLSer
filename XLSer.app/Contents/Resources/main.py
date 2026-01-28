try:
    import tkinter as tk
    from tkinter import filedialog
    import openpyxl
    from openpyxl import Workbook


    # 2 часа
    # 2 часа

    # Загрузка папки §
    # Конвертация в xlsx §
    # Конвертация в numbers

    def UploadAction(event=None):
        # Загрузка файла
        filenames = filedialog.askopenfilenames()
        root.tk.splitlist(filenames)
        print('Путь выбанного файла:', filenames)
        Convert_data(filenames)


    def Convert_data(filenames):
        result_list = [[" ", "Фактический остаток", "Книжный остаток"]]
        result_list2 = [["№ п/п", "№ АЗС / МАЗС / ААЗС",
                         "GD-95 излишки (л.)", "GD-95 недостача (л.)",
                         "Аи-92 излишки (л.)", "Аи-92 недостача (л.)",
                         "Аи-95 излишки (л.)", "Аи-95 недостача (л.)",
                         "Диз. топ. излишки (л.)", "Диз. топ. недостача (л.)",
                         "СУГ излишки (л.)", "СУГ недостача (л.)",
                         "всего излишки", "всего недостачи"
                         ]]
        station_counter=0
        for filename in filenames:

            wb = openpyxl.load_workbook(filename=filename)
            ws = wb['ssoMain']

            # находим конец основной таблицы
            target_value = 'Таб. 2. Движение СТиУ'
            a = 0
            target_found = False
            column_to_check = 2  # Column A (openpyxl uses 1-based indexing)

            # Iterate through rows up to the maximum row that might contain data
            for row_num in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row_num, column=column_to_check).value

                if cell_value == target_value:
                    a = row_num - 1  # Count of rows *until* the target cell
                    target_found = True
                    break  # Stop the loop when the cell is found

                # If counting total rows including the target cell itself,
                # you would use `row_count = row_num`

            if target_found:
                print(f"The cell with value '{target_value}' was found at row {row_num}.")
                print(f"There are {a} rows before it in Column {column_to_check}.")
            else:
                print(f"The target value '{target_value}' was not found in the column.")




            # Читаем нужные нам столбцы
            # Столбец НП
            col_B = [cell.value for cell in ws['B'][1:]]
            print(col_B)
            # Столбец Резервуар
            col_D = [cell.value for cell in ws['D'][1:]]
            print(col_D)
            # Столбец Факт.+трубопр., л/кг
            # фактические остатки
            col_fact = [cell.value for cell in ws['BP'][1:]]
            print(col_fact)
            # Столбец Книжный, л / кг
            # книжные остатки
            col_book = [cell.value for cell in ws['BR'][1:]]
            print(col_book)

            # Читаем столбцы для излишков и недостач
            # Излишки
            col_surplus = [cell.value for cell in ws['BV'][1:]]
            print(f"Излишки: {col_surplus}")
            # Столбец Резервуар
            col_shortage = [cell.value for cell in ws['BZ'][1:]]
            print(f"Недосатчи: {col_shortage}")

            gas_type = ["GD95", "Аи-92", "Аи-95", "ДТ", "СУГ", "???"]

            intindex = a


            current_type = gas_type[0]
            # result_list = [[" ", "Фактический остаток", "Книжный остаток"]]


            lastslash = filename.rfind('/') + 1
            station_name = filename[lastslash:-5]#номер заправки
            result_list.append([station_name, "", ""])
            for i in range(8, intindex):
                if col_B[i] in gas_type:
                    current_type = col_B[i]
                    for j in range(i, intindex):
                        if col_B[j] == "Итого":
                            result_list.append([current_type, col_fact[j], col_book[j]])
                            break

            station_counter += 1
            current_surplus = 0
            current_shortage = 0

            result_list2.append([station_counter, station_name,
                                 "","",
                                 "","",
                                 "","",
                                 "","",
                                 "", "",
                                 "",""])


            for i in range(8, intindex):
                if col_B[i] in gas_type:
                    current_type = col_B[i]
                    for j in range(i, intindex):

                        if col_B[j] == "Итого":
                            # Cчитаем сумму
                            if col_surplus[j] is not None:
                                current_surplus += float(col_surplus[j])
                            if col_shortage[j] is not None:
                                current_shortage += float(col_shortage[j])
                            # Заносим отдельные показатели в таблицу
                            if current_type == gas_type[0]:
                                result_list2[station_counter][2] = col_surplus[j]
                                result_list2[station_counter][3] = col_shortage[j]
                            elif current_type == gas_type[1]:
                                result_list2[station_counter][4] = col_surplus[j]
                                result_list2[station_counter][5] = col_shortage[j]
                            elif current_type == gas_type[2]:
                                result_list2[station_counter][6] = col_surplus[j]
                                result_list2[station_counter][7] = col_shortage[j]
                            elif current_type == gas_type[3]:
                                result_list2[station_counter][8] = col_surplus[j]
                                result_list2[station_counter][9] = col_shortage[j]
                            elif current_type == gas_type[4]:
                                result_list2[station_counter][10] = col_surplus[j]
                                result_list2[station_counter][11] = col_shortage[j]
                            break

            result_list2[station_counter][12] = current_surplus
            result_list2[station_counter][13] = current_shortage

            res_wb = Workbook()
            res_ws = res_wb.active
            for row in result_list:
                res_ws.append(row)

            file_path = "output.xlsx"
            res_wb.save(file_path)
            if filename == filenames[-1]:
                Save_res(res_wb)

            res_wb2 = Workbook()
            res_ws2 = res_wb2.active
            for row in result_list2:
                res_ws2.append(row)

            file_path2 = "output2.xlsx"
            res_wb2.save(file_path2)
            if filename == filenames[-1]:
                Save_res(res_wb2)



    def Save_res(res_df):

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")],
            title="Сохранить как..."
        )

        # Если пользователь не отменил диалог
        if filepath:
            res_df.save(filepath)
            print(f"Файл сохранён: {filepath}")


    # Начальное окно
    root = tk.Tk()
    root.geometry("350x200")
    root.title('Книжные и факт остатки')
    # Кнопка для загрузки файла
    button = tk.Button(root, text='Загрузить файл', command=UploadAction)
    button.pack(padx=6, pady=24)

    root.mainloop()
except:
    print("An exception occurred")
