import tkinter as tk
import os
from tkinter import messagebox

""""
При запуску програми буде запускатися файл ексель з початковими розрахунками
"""


# os.startfile("C:\\Users\\User\\Desktop\\for_make_protocol\\dum_toxic\\Toksik_dum.xlsx")

class FrameButton:

    def __init__(self,
                 rowFrame,
                 root,
                 dictparam
                 ):
        self.rowFrame = rowFrame
        self.root = root
        self.dictparam = dictparam

    def open_file(self, index):
        os.startfile(self.dictparam['list_with_path'][index])

    def button_on_tk(self):
        # ------------------- Рамка з кнопкам идля первинних протоолів ----------
        lbl_frm_pervini_protocol = tk.LabelFrame(self.root, text=self.dictparam['listMane_Frame_and_button'][0],
                                                 font='15')
        lbl_frm_pervini_protocol.grid(row=self.rowFrame, column=1, pady=10, padx=30)

        btn_1 = tk.Button(lbl_frm_pervini_protocol,
                          text=self.dictparam['listMane_Frame_and_button'][1], font='15',
                          width=20, height=3, bg=self.dictparam['list_Colors_botton'][0],
                          relief="raised", command=lambda: self.open_file(0))
        btn_1.grid(row=1, column=1, pady=10, padx=30)

        btn_2 = tk.Button(lbl_frm_pervini_protocol,
                          text=self.dictparam['listMane_Frame_and_button'][2], font='15',
                          width=20, height=3, bg=self.dictparam['list_Colors_botton'][1],
                          relief="raised", command=lambda: self.open_file(1))
        btn_2.grid(row=1, column=2, pady=10, padx=30)

        btn_3 = tk.Button(lbl_frm_pervini_protocol,
                          text=self.dictparam['listMane_Frame_and_button'][3], font='15',
                          width=20, height=3, bg=self.dictparam['list_Colors_botton'][2],
                          relief="raised", command=lambda: self.open_file(2))
        btn_3.grid(row=1, column=3, pady=10, padx=30)


class Renamefiles:

    def __init__(self,
                 root,
                 rowFrame,

                 ):
        self.root = root
        self.rowFrame = rowFrame

    def build_label_and_btn_Renam(self):
        var_entry_path_file = tk.StringVar()
        var_entry_path_file.set(' ')


        lbl_frm_rename_protocols = tk.LabelFrame(self.root, text='ПЕРЕЙМЕНУВАТИ_ФАЙЛИ',
                                                 borderwidth=1, font=('Arial, 12 ' ),
                                                 bg='#B2DFEE')
        lbl_frm_rename_protocols.grid(row=self.rowFrame, column=1, pady=10, padx=30)

        lbl_path_to_id_file = tk.Label(lbl_frm_rename_protocols,
                                       text='ВКАЖІТЬ ШЛЯХ ДО ПАПКИ\n З ФАЙЛОМ ІДЕНТИФІКАЦІЇ',
                                       font=('Arial, 12 ' )
        )
        lbl_path_to_id_file.grid(row=1, column=1, padx=5, pady=10)

        path_to_id_file = tk.Entry(lbl_frm_rename_protocols, width=41, font='Arial, 14',
                                   )
        path_to_id_file.grid(row=1, column=2, columnspan=1, padx=5, pady=10)


        def event_entry_path_for_rename_files(event):

            """"
            потрібно, щоб назва папки  якій знаходяться файли відповідала назві майбутнього файлу.
             Тобто користувач перейменовує папку, а файли в ній НІ.

            """""
            try:
                " будемо видаляти зайвий(останній) слеш з шляху до файлу "
                if path_to_id_file.get()[-1] == '\\':
                    path = path_to_id_file.get()[0:-1]
                else:
                    path = path_to_id_file.get()

                original_name_dir = path.split(f"\\")[-1]
                split_path_dir = original_name_dir.split('_')
                new_name_file_idetif = split_path_dir[0] + '_ПІ_' + '_'.join(split_path_dir[1:-1])

                new_name_file = '_'.join(split_path_dir[0:-1])

                standart_path = path + '\\'  # стала частина шляху до файлу

                for name_file in os.listdir(path):
                    #  перевіряємо чи в папці є інші папки аби з ними нічого не зробити

                    try:
                        type_file = name_file.split('.')[1]
                        if 'doc' in name_file:
                            #  перейменовуємо файли з потрібним розширенням
                            try:
                                if 'ПІ' in name_file:
                                    new_name_idetif = new_name_file_idetif + '.' + type_file
                                    os.rename(standart_path + name_file, standart_path + new_name_idetif)
                                else:
                                    new_name_protokol_file = new_name_file + '.' + type_file
                                    os.rename(standart_path + name_file, standart_path + new_name_protokol_file)

                            except Exception:
                                print('Exception_remove_not_doc_file ')

                        else:
                            path_to_id_file.delete(0, 'end')  # Видаляє з поля entry вміст і робить його чистим
                            os.remove(standart_path + name_file)  # видаляємо файли, що не відповідають розширенню


                    except Exception:
                        pass

            except Exception:

                tk.messagebox.showinfo(title='!!Помилка!!!', message='ПЕРЕВІРТЕ ПРАВИЛЬНІСТЬ ВЕДЕННЯ ШЛЯХУ ДО ФАЙЛІВ')
                path_to_id_file.delete(0, 'end')  # Видаляє з поля entry вміст і робить його чистим

        " Подія, шр має перйменовувати файли пісял натиску кнопки ентер"
        path_to_id_file.bind('<Return>', event_entry_path_for_rename_files)


        "************ ПЕРЕВЕДЕННЯ KW->HK"
        var_kw = tk.StringVar()  # змінна для маніпуляції
        var_kw.set('')
        lbl_frm_transwer_kw = tk.LabelFrame(lbl_frm_rename_protocols, text="ПЕРЕВЕДЕННЯ KW->HK ", width=30, height=20)
        lbl_frm_transwer_kw.grid(row=1, column=3, pady=5, padx=10)

        entry_field_kw = tk.Entry(lbl_frm_transwer_kw, width=5, font='Arial, 14', textvariable=var_kw)
        entry_field_kw.grid(row=1, column=1, pady=10, padx=30)

        lbl_with_HK = tk.Label(lbl_frm_transwer_kw, text=f'HK =', width=15, height=3, font='15',
                               borderwidth=3, bg='green')
        lbl_with_HK.grid(row=2, column=1, pady=10, padx=10)

    # ----
        def event_transfer_kw_to_hk(event):
            lbl_with_HK['text'] = f'HK ={round(int(var_kw.get()) * 1.36, 2)}'

        # Застосовуємо подiю
        entry_field_kw.bind('<Return>', event_transfer_kw_to_hk)


    #
    # def entry_path_for_rename_files(self, path_to_id_file):
    #
    #     """"
    #     потрібно, щоб назва папки  якій знаходяться файли відповідала назві майбутнього файлу.
    #      Тобто користувач перейменовує папку, а файли в ній НІ.
    #
    #     """""
    #     try:
    #         path = path_to_id_file.get()
    #
    #         original_name_dir = path.split(f"\\")[-1]
    #         split_path_dir = original_name_dir.split('_')
    #         new_name_file_idetif = split_path_dir[0] + '_ПІ_' + '_'.join(split_path_dir[1:-1])
    #
    #         new_name_file = '_'.join(split_path_dir[0:-1])
    #         # print(
    #         #     'new_name_file_idetif = ', new_name_file_idetif,
    #         #     '\nnew_name_file = ', new_name_file
    #         # )
    #
    #         standart_path = path + '\\'  # стала частина шляху до файлу
    #         for name_file in os.listdir(path):
    #             #  перевіряємо чи в папці є інші папки аби з ними нічого не зробити
    #             try:
    #                 type_file = name_file.split('.')[1]
    #                 if 'doc' not in name_file:
    #                     os.remove(standart_path + name_file)  # видаляємо файли, що не відповідають розширенню
    #             except Exception:
    #                 pass
    #
    #             #  перейменовуємо файли з потрібним розширенням
    #             try:
    #                 if 'ПІ' in name_file:
    #                     new_name_idetif = new_name_file_idetif + '.' + type_file
    #                     os.rename(standart_path + name_file, standart_path + new_name_idetif)
    #                 else:
    #                     new_name_protokol_file = new_name_file + '.' + type_file
    #                     os.rename(standart_path + name_file, standart_path + new_name_protokol_file)
    #
    #             except Exception:
    #                 pass
    #
    #         path_to_id_file.delete(0, 'end')  # Видаляє з поля entry вміст і робить його чистим
    #     except Exception:
    #         tk.messagebox.showinfo(title='!!Помилка!!!', message='ПЕРЕВІРТЕ ПРАВИЛЬНІСТЬ ВЕДЕННЯ ШЛЯХУ ДО ФАЙЛІВ')
    #
    #         path_to_id_file.delete(0, 'end')  # Видаляє з поля entry вміст і робить його чистим

    # # видалення всіх фалів з папки
    # # спочатку скануємо вміст папки, а потім видаляємо те що там є через задання шляху до відповіного файлу


# class KwatIn_hk:
#     def __init__(self,
#                  rowFrame,
#                  root
#                  ):
#         self.rowFrame = rowFrame
#         self.root = root
#
#     def frame_where_transwer_kw(self):
#         var_kw = tk.StringVar()  # змінна для маніпуляції
#         var_kw.set('')
#         lbl_frm_transwer_kw = tk.LabelFrame(self.root, text="ПЕРЕВЕДЕННЯ KW->HK ", width=30, height=20)
#         lbl_frm_transwer_kw.grid(row=self.rowFrame, column=3, pady=10, padx=30)
#
#         entry_field_kw = tk.Entry(lbl_frm_transwer_kw, width=10, font='12', textvariable=var_kw)
#         entry_field_kw.grid(row=1, column=1, pady=10, padx=30)
#
#         lbl_with_HK = tk.Label(lbl_frm_transwer_kw, text=f'HK =', width=15, height=3, font='15',
#                                borderwidth=3, bg='green')
#         lbl_with_HK.grid(row=2, column=1, pady=10, padx=30)
#
#         # ----
#         def event_transfer_kw_to_hk(event):
#             lbl_with_HK['text'] = f'HK ={round(int(var_kw.get()) * 1.36, 2)}'
#
#         # Застосовуємо подiю
#         entry_field_kw.bind('<Return>', event_transfer_kw_to_hk)


root = tk.Tk()
root.title("ПОМІЧНИК З СТВОРЕННЯ ПРОТОКОЛІВ ВИПРОБУВАНЬ")
# root.geometry('690x300')
w, h = 100, 80
root.minsize(width=w, height=h)

dictparam_pp_M1_N1 = {
    'listMane_Frame_and_button': ['ПЕРВИННІ_ПРОТОКОЛИ',
                                  'М1',
                                  'М1_67',
                                  'N1'
                                  ],
    'list_Colors_botton': ['#FFFAF0',
                           '#FFFAF0',
                           '#F8F8FF'
                           ],
    'list_with_path': ['.\\for_make_protocol\\M1\\1_M1_PP_with_date_and_identif.docx',
                       '.\\for_make_protocol\\M1\\2_M1_PP_R67_with_date_and_identif.docx',
                       '.\\for_make_protocol\\N1\\1_N1_PP_with_date_and_identif.docx'
                       ]
}

dictparam_vikidu = {
    'listMane_Frame_and_button': ['ТОКСИЧНІСТЬ ТА ВИКИДИ',
                                  'БЕНЗИН, ГАЗ',
                                  'ДИЗЕЛЬ',
                                  'Ексель файл \n з Гальмами та Викидами'
                                  ],
    'list_Colors_botton': ['#CAFF70',
                           '#BCEE68',
                           '#A2CD5A'
                           ],
    'list_with_path': ['.\\for_make_protocol\\dum_toxic\\2_Gas_Catalyzed_with_lambda_probe.docx',
                       '.\\for_make_protocol\\dum_toxic\\1_Turbo_Diesel_.docx',
                       '.\\for_make_protocol\\dum_toxic\\Toksik_dum.xlsx'
                       ]
}

if __name__ == '__main__':
    obj_1 = FrameButton(rowFrame=1, root=root, dictparam=dictparam_pp_M1_N1)
    obj_1.button_on_tk()

    obj_2 = FrameButton(rowFrame=2, root=root, dictparam=dictparam_vikidu)
    obj_2.button_on_tk()

    obj_3_renames_files = Renamefiles(rowFrame=3, root=root)
    obj_3_renames_files.build_label_and_btn_Renam()

    # obj_4_transwer_power = KwatIn_hk(rowFrame=3, root=root)
    # obj_4_transwer_power.frame_where_transwer_kw()

    tk.mainloop()
