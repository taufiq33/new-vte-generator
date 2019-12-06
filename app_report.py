import xlrd
import xlsxwriter
import string
import math
import datetime
import tkinter
import tkinter.filedialog
import app


class TemplateVoucher():
    cols_alpha = {
        1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M',
        14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'
    }
    penomoran = list(range(1, 51))

    def __init__(self, lokasi_file):
        self.namafile = lokasi_file
        self.objek_file = xlrd.open_workbook(lokasi_file)
        self.sheet_file = self.objek_file.sheet_by_index(0)

        self.kode_voucher = self.get_kode_voucher()
        self.jumlah_voucher = self.get_jumlah_voucher()
        self.tanggal = self.get_tanggal_voucher()
        self.kodegosok = self.get_kode_gosok()

        self.kolom = self.get_total_kolom()

    def get_total_kolom(self):
        kol = 2
        if self.jumlah_voucher > 50:
            kol = math.ceil(self.jumlah_voucher / 50) + 1
        return kol

    def get_kode_voucher(self):
        return self.sheet_file.cell_value(0, 0)

    def get_jumlah_voucher(self):
        self.sheet_file.cell_value(0, 1)
        return self.sheet_file.nrows

    def get_tanggal_voucher(self):
        tgl = self.namafile.split(" ")[-4]
        return tgl.replace("-", "_")

    def get_kode_gosok(self):
        data_kode_gosok = []
        for kode_gosok in range(0, self.jumlah_voucher):
            data_kode_gosok.append(self.sheet_file.cell_value(kode_gosok, 1))
        return data_kode_gosok

    def __str__(self):
        return "{} {} - {}pcs".format(self.kode_voucher, self.tanggal, self.jumlah_voucher)


class FileTemplateVoucher():
    def __init__(self, arrayTemplateVoucher, sift):
        self.sift = sift  # pagi , sore, malam/dinihari
        self.array_template_voucher = arrayTemplateVoucher
        self.file_title = "REPORT ketikan {} - SIFT {}.xlsx".format(
            str(datetime.datetime.now()).replace(":", "_")[:19],
            self.sift
        )
        self.file_excel = xlsxwriter.Workbook(
            app.getPathConfig() + "report/" + self.file_title)
        self.worksheet = self.file_excel.add_worksheet()

        self.total_kolom = self.sumKolom()

        self.write_data()

    def sumKolom(self):
        k = 0
        for template in self.array_template_voucher:
            k += template.kolom
        return k

    def write_data(self):
        kolom = 0
        baris = 0
        alphabet = 1
        maks_kolom = 19
        count_kolom = 0
        
        
        formattext = self.file_excel.add_format()
        formattext.set_num_format('@')  # == text in excel

        kuning = self.file_excel.add_format({'bg_color':'#FFFF00'})
        kuning.set_num_format('@')
        for template in self.array_template_voucher:
            
            self.worksheet.write(
                baris, kolom, template.kode_voucher + " " + self.sift, formattext)
            self.worksheet.write(
                baris + 1, kolom, template.tanggal, formattext)
            x = baris + 2
            for penomoran in template.penomoran:
                self.worksheet.write(x, kolom, penomoran, formattext)
                x += 1
            y = 1
            for x in range(2, template.kolom + 1):
                self.worksheet.write(
                    baris, kolom + y, template.cols_alpha[x - 1], formattext)
                y += 1

            x = baris + 2
            y = 1
            if template.jumlah_voucher <= 50:
                for kodegosok in template.kodegosok:
                    self.worksheet.write(x, kolom + y, kodegosok, kuning)
                    x += 1
            else:
                count = 1
                for kodegosok in template.kodegosok:
                    self.worksheet.write(x, kolom + y, kodegosok, kuning)
                    x += 1
                    count += 1
                    if count > 50:
                        y += 1
                        count = 1
                        x = baris + 2
            kolom = kolom + template.kolom + 1
            count_kolom = count_kolom + template.kolom

            if count_kolom >= maks_kolom :
                baris = baris + 55
                count_kolom = 0
                kolom = 0

        self.worksheet.set_column('A1:ZZ50', 25)
        self.file_excel.close()
        return True

# root = tkinter.Tk()
# root.title("coba")
# filez = tkinter.filedialog.askopenfilenames(parent=root, title='Choose a file', )

# obj = []
# for x in filez:
#     obj.append(TemplateVoucher(x))


# ini = FileTemplateVoucher(obj, 'pagi')
# ini.write_data()

# root.mainloop()
