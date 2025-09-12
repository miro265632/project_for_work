import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import csv


class ExcelConverterApp:
    def __init__(self, master):
        self.master = master
        master.title("Конвертер XLS в CSV")

        self.label = tk.Label(master, text="Выберите XLS файл для обработки:")
        self.label.pack()

        self.select_button = tk.Button(master, text="Выбрать файл", command=self.select_file)
        self.select_button.pack()

        self.convert_button = tk.Button(master, text="Преобразовать и скачать", command=self.convert_file,
                                        state=tk.DISABLED)
        self.convert_button.pack()

        self.file_path = ""

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if self.file_path:
            self.convert_button.config(state=tk.NORMAL)

    def convert_file(self):
        try:
            # Чтение Excel файла  
            df = pd.read_excel(self.file_path)
            # Если файл имеет нужные столбцы, продолжите с обработкой  
            # Здесь вам нужно будет адаптировать вашу логику обработки  
            a = df.values.tolist()

            c = a[0]
            q = []
            k=[]
            r = a.copy()
            uid = []
            for i in r:
                uid.append(i[0])
                del i[0]
                del i[-1]
                del i[-1]
                del i[-2]
                x = int(i[0].split()[-1][:-4])
                print(x)
                y = int(i[1])
                z = y / x
                z = round(z)
                print(z)
                m=x*z
                print(m)
                q.append(z)
                k.append(m)

            for i in a:
                for j in range(len(q)):
                    i.append(str(q[j]))
                    del q[j]
                    break
            for i in a:
                for j in range(len(k)):
                    i.append(str(k[j]))
                    del k[j]
                    break
            for i in a:
                for j in range(len(uid)):
                    i.insert(0, str(uid[j]))
                    del uid[j]
                    break

            a.insert(0, ['UUID', 'Наименование', 'Количество', 'Количество упаковок', 'Округление количество товаров'])

            # Сохранение результата в новый Excel файл  
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                       filetypes=[("Excel files", "*.xlsx;*.xls")])
            if not output_file:
                return  # Если пользователь закрыл диалог, ничего не делаем
            df_output = pd.DataFrame(a[1:], columns=a[0])
            df_output.to_excel(output_file, index=False)

            # Вывод сообщения об успешном завершении  
            messagebox.showinfo("Успех", f"Файл успешно преобразован и сохранён как {output_file}")

        except Exception as e:
            messagebox.showerror("Ошибка", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelConverterApp(root)
    root.mainloop()