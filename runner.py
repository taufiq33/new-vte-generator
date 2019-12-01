from tkinter import *
import tkinter.messagebox
import tkinter.filedialog
import app
import os
fields = ('Kode voucher', 'Deskripsi', 'Jumlah')


def pengaturan_event(e):
    newwin = Toplevel(root)
    newwin.title("Pengaturan")
    img = tkinter.PhotoImage(file="setting.png")
    newwin.tk.call('wm', 'iconphoto', newwin._w, img)
    newwin.geometry("600x200+300+300")

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

def simpan_event(e):
    folder = e[0].get()
    kadaluarsa = e[1].get()
    if app.updateConfig(kadaluarsa, folder) :
        tkinter.messagebox.showinfo("Berhasil", "Berhasil simpan pengaturan, silahkan buka kembali aplikasi.")
        root.quit()
    
def pilih_event(z):
    dirname = tkinter.filedialog.askdirectory()
    if dirname :
        dirname += '/'
        z.delete(0, 'end')
        z.insert(0, dirname)

def generate_file_event(entries):
    kd = entries['Kode voucher'].get()
    desk = entries['Deskripsi'].get()
    jumlah = entries['Jumlah'].get()
    retail = False

    if kd is "" or desk is "" or jumlah is "":
        return tkinter.messagebox.showerror("Gagal", "jangan ada yang kosong ya om")

    if CheckVar1.get() == 1:
        kd += ",OTOMAX RETAIL"
        retail = True
    # print(kd, desk, jumlah)
    fileExcel = app.ExcelFile(kd, desk, jumlah, retail)
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
   c1 = Checkbutton(root, text="Otomax Retail ?", variable=CheckVar1,
                    onvalue=1, offvalue=0, height=5,
                    width=20)
   c1.pack(side=LEFT, padx=5, pady=5)
   b1 = Button(root, text='Pengaturan',
               command=(lambda e=ents: pengaturan_event(e)))
   b1.pack(side=LEFT, padx=5, pady=5)
   b2 = Button(root, text='Generate File',
               command=(lambda e=ents: generate_file_event(e)))
   b2.pack(side=LEFT, padx=5, pady=5)
   b3 = Button(root, text='Quit', command=root.quit)
   b3.pack(side=LEFT, padx=5, pady=5)
   root.mainloop()
