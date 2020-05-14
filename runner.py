from tkinter import *
import tkinter.messagebox
import tkinter.filedialog
import app
import app_report
import os
fields = ('Kode voucher', 'Deskripsi', 'Jumlah')


def pengaturan_event(e):
    newwin = Toplevel(root)
    newwin.title("Pengaturan")
    img = tkinter.PhotoImage(file="setting.png")
    newwin.tk.call('wm', 'iconphoto', newwin._w, img)
    newwin.geometry("600x200+100+100")
    # newwin.attributes('-topmost', 'true')

    entries = {}
    row1 = Frame(newwin)
    row2 = Frame(newwin)
    row3 = Frame(newwin)
    lab_folder = Label(row2, width=22, text="Folder: ", anchor='w')
    lab_kadaluarsa = Label(row1, width=22, text="Kadaluarsa: ", anchor='w')
    ent_folder = Entry(row2)
    ent_kadaluarsa = Entry(row1)
    btn_pilih = Button(row3, text="Select Folder ...",
                       command=(lambda e=ent_folder: pilih_event(e)))

    ent_kadaluarsa.insert(0, app.getKadaluarsaConfig())
    ent_folder.insert(0, app.getPathConfig())

    row1.pack(side=TOP, fill=X, padx=5, pady=5)
    row2.pack(side=TOP, fill=X, padx=5, pady=5)
    row3.pack(side=TOP, fill=X, padx=5, pady=5)
    lab_kadaluarsa.pack(side=LEFT)
    ent_kadaluarsa.pack(side=RIGHT, expand=YES, fill=X)
    lab_folder.pack(side=LEFT)
    ent_folder.pack(side=RIGHT, expand=YES, fill=X)
    btn_pilih.pack(side=RIGHT, padx=5, pady=5)

    b1 = Button(newwin, text='Batal',
                command=newwin.destroy)
    b1.pack(side=RIGHT, padx=5, pady=5)
    b2 = Button(newwin, text='Simpan Pengaturan',
                command=(lambda e=[ent_folder, ent_kadaluarsa]: simpan_event(e)))
    b2.pack(side=RIGHT, padx=5, pady=5)


def report_event(e):
    reportwin = Toplevel(root)
    reportwin.title("Buat Report")
    reportwin.geometry("600x600+700+700")
    # reportwin.attributes('-topmost', 'true')

    tkinter.messagebox.showwarning(
        "Perhatian", "Pastikan anda sudah membuat folder 'report' didalam folder {}".format(app.getPathConfig()))

    row0 = Frame(reportwin)
    row1 = Frame(reportwin)
    row2 = Frame(reportwin)
    files_text = Text(row2,)
    lab_sift = Label(row0, width=22, text="Sift: ", anchor='w')
    ent_sift = Entry(row0)
    # files_text.config(state=DISABLED)
    btn_pilih = Button(row1, text="Select Files ...",
                       command=(lambda e=files_text: pilih_files_event(e)))

    row0.pack(side=TOP, fill=X, padx=5, pady=5)
    row1.pack(side=TOP, fill=X, padx=5, pady=5)
    row2.pack(side=TOP, fill=X, padx=5, pady=5)
    lab_sift.pack(side=LEFT)
    ent_sift.pack(side=RIGHT, expand=YES, fill=X)
    files_text.pack(side=LEFT)
    btn_pilih.pack(side=LEFT, padx=5, pady=5)

    b1 = Button(reportwin, text='Create Report >> ',
                command=(lambda e=[files_text, ent_sift]: create_report_event(e)))
    b1.pack(side=RIGHT, padx=5, pady=5)
    b2 = Button(reportwin, text='Batal',
                command=reportwin.destroy)
    b2.pack(side=RIGHT, padx=5, pady=5)


def simpan_event(e):
    folder = e[0].get()
    kadaluarsa = e[1].get()
    if app.updateConfig(kadaluarsa, folder):
        tkinter.messagebox.showinfo(
            "Berhasil", "Berhasil simpan pengaturan, silahkan buka kembali aplikasi.")
        root.quit()


def create_report_event(e):
    temp = e[0].get("1.0", END)
    sift = e[1].get()

    if temp == '' or sift == '':
        return tkinter.messagebox.showerror("Gagal", "jangan ada yang kosong ya om")

    temp = temp.split("\n")
    temp.remove("")

    obj = []
    for x in temp:
        obj.append(app_report.TemplateVoucher(x))

    report = app_report.FileTemplateVoucher(obj, sift)
    if report.write_data():
        return tkinter.messagebox.showinfo("Succes", "Berhasil > {}".format(report.file_title))


def pilih_event(z):
    dirname = tkinter.filedialog.askdirectory()
    if dirname:
        dirname += '/'
        z.delete(0, 'end')
        z.insert(0, dirname)


def pilih_files_event(e):
    filez = tkinter.filedialog.askopenfilenames(title='Choose files')
    e.insert(INSERT, "\n".join(filez))
    e.state = "DISABLED"

def about_event(e):
    isi_pesan = """Voucher to Excel Generator v0.5\n\nSoftware sederhana untuk membantu proses input voucher ke SW Pulsa.\n\nPembuat : Taufiq Hidayat\nRepo github : github.com/taufiq33/new-vte-generator"""
    return tkinter.messagebox.showinfo("About this software",isi_pesan.strip())


def generate_file_event(entries):
    kd = entries['Kode voucher'].get()
    desk = entries['Deskripsi'].get()
    jumlah = entries['Jumlah'].get()
    retail = False
    hint2 = False
    hint3 = False

    if kd is "" or desk is "" or jumlah is "":
        return tkinter.messagebox.showerror("Gagal", "jangan ada yang kosong ya om")

    if CheckVar1.get() == 1:
        kd += ",OTOMAX RETAIL"
        retail = True
    if CheckVar2.get() == 1:
        hint2 = True
    if CheckVar3.get() == 1:
        hint3 = True
    # print(kd, desk, jumlah)
    fileExcel = app.ExcelFile(kd, desk, jumlah, retail, hint2, hint3)
    if fileExcel.generateExcelFile():
        entries['Kode voucher'].delete(0, 'end')
        entries['Deskripsi'].delete(0, 'end')
        entries['Jumlah'].delete(0, 'end')
        tkinter.messagebox.showinfo("Success", "file berhasil dibuat")


def makeform(root, fields):
   entries = {}
   for field in fields:
      row = Frame(root)
      lab = Label(row, width=22, text=field+": ", anchor='w')
      ent = Entry(row)
      row.pack(side=TOP, fill=X, padx=5, pady=5)
      lab.pack(side=LEFT)
      ent.pack(side=RIGHT, expand=YES, fill=X)
      entries[field] = ent
   return entries


if __name__ == '__main__':
   root = Tk()
   root.title("VTE - Voucher to Excel Generator")
   img = tkinter.PhotoImage(file="icon.png")
   root.tk.call('wm', 'iconphoto', root._w, img)
   ents = makeform(root, fields)
   root.bind('<Return>', (lambda event, e=ents: fetch(e)))
   CheckVar1 = IntVar()
   c1 = Checkbutton(root, text="Retail ?", variable=CheckVar1,
                    onvalue=1, offvalue=0, height=5,
                    width=20)
   c1.pack(side=LEFT, padx=5, pady=5)
   CheckVar2 = IntVar()
   c2 = Checkbutton(root, text="Hint kelipatan2?", variable=CheckVar2,
                    onvalue=1, offvalue=0, height=5,
                    width=20)
   c2.pack(side=LEFT, padx=5, pady=5)
   CheckVar3 = IntVar()
   c3 = Checkbutton(root, text="Hint kelipatan3?", variable=CheckVar3,
                    onvalue=1, offvalue=0, height=5,
                    width=20)
   c3.pack(side=LEFT, padx=5, pady=5)
#    b1 = Button(root, text='Pengaturan',
#                command=(lambda e=ents: pengaturan_event(e)))
#    b1.pack(side=LEFT, padx=5, pady=5)
#    btnReport = Button(root, text="Buat Report",
#                       command=(lambda e=ents: report_event(e))
#                       )
#    btnReport.pack(side=LEFT, padx=5, pady=5)
   b2 = Button(root, text='Generate File',
               command=(lambda e=ents: generate_file_event(e)))
   b2.pack(side=LEFT, padx=5, pady=5)
#    b3 = Button(root, text='Quit', command=root.quit)
#    b3.pack(side=LEFT, padx=5, pady=5)

   menubar = Menu(root)
   optionsmenu = Menu(menubar, tearoff = 0)
   optionsmenu.add_separator()
   menubar.add_cascade(label="Options", menu = optionsmenu)
   optionsmenu.add_command(label="Preferences", command=(lambda e=ents: pengaturan_event(e)))

   reportmenu = Menu(menubar, tearoff=0)
   reportmenu.add_separator()
   menubar.add_cascade(label="Report", menu = reportmenu)
   reportmenu.add_command(label="Create Report", command=(lambda e=ents: report_event(e)))
   
   helpmenu = Menu(menubar, tearoff=0)
   helpmenu.add_separator()
   menubar.add_cascade(label="Help", menu = helpmenu)
   helpmenu.add_command(label="About", command=(lambda e=ents: about_event(e)))
   
   root.config(menu = menubar)
   root.mainloop()
