import pandas as pd
import xlwt

#variabel array menyimpan data berbentuk maps dari excel
data_mhs = []

#variabel batas
#batas penghasilan
batasA = 5.8
batasB = 6.48
batasC = 7.1
batasD = 12.2

#batas pengeluaran
batasLuarA = 4
batasLuarB = 5.1
batasLuarC = 6
batasLuarD = 8.7

#variabel array untuk MAIN PROGRAM
luar, hasil, inf, yStar, index = [], [], [], [], []

#fungsi untuk read data excel (sebelum dijalankan pindahkan path folder pada terminal ke folder program yang berisi file excel)
def readData() :
    df = pd.read_excel('Mahasiswa.xls', sheet_name='Mahasiswa')

    #data dari excel disimpan dalam bentuk maps dengan key dan value
    for i in df.index:
        data = {}
        data['id'] = df['Id'][i]
        data['penghasilan'] = df['Penghasilan'][i]
        data['pengeluaran'] = df['Pengeluaran'][i]
        data_mhs.append(data)
        
# RUMUS LINIER NAIK DAN TURUN

def naik(val, a, b) :
    return (val - a) / (b - a)

def turun(val, a, b) :
    return (-1 * (val - b))/(b - a)

#FUZZIFIKASI PENGHASILAN DAN PENGELUARAN
def penghasilan(value, a, b, c, d) :
    kecil, menengah, besar = 0, 0, 0

    if (value <= a) :
        kecil = 1
    elif (value > b) :
        kecil = 0
    elif (value > a and value <= b) :
        kecil = turun(value, a, b)
    
    if (value <= a and value > d) :
        menengah = 0
    elif (value > a and value <= b) :
        menengah = naik(value, a, b)
    elif (value > b and  value <= c) :
        menengah = 1
    elif (value > c and value <= d) :
        menengah = turun(value, c, d)
    
    if (value <= c) :
        besar = 0
    elif (value > d) :
        besar = 1
    elif (value > c and value <= d) :
        besar = naik(value, c, d)
    
    return [kecil, menengah, besar]

def pengeluaran(value, a, b, c, d) :
    sedikit, cukup, banyak = 0, 0, 0

    if (value <= a) :
        sedikit = 1
    elif (value > b) :
        sedikit = 0
    elif (value > a and value <= b) :
        sedikit = turun(value, a, b)
    
    if (value <= a and value > d) :
        cukup = 0
    elif (value > a and value <= b) :
        cukup = naik(value, a, b)
    elif (value > b and  value <= c) :
        cukup = 1
    elif (value > c and value <= d) :
        cukup = turun(value, c, d)
    
    if (value <= c) :
        banyak = 0
    elif (value > d) :
        banyak = 1
    elif (value > c and value <= d) :
        banyak = naik(value, c, d)
    
    return [sedikit, cukup, banyak]

#INFERENSI MENERAPKAN SUGENO RULE
def inferensi(hasil, keluar) :
    tidak, tengah, layak = [], [], []

    tidak.append(max(hasil[1], keluar[0]))
    tidak.append(max(hasil[2], keluar[0]))
    tidak.append(max(hasil[2], keluar[1]))

    tengah.append(max(hasil[0], keluar[0]))
    tengah.append(max(hasil[1], keluar[1]))
    tengah.append(max(hasil[2], keluar[2]))

    layak.append(max(hasil[0], keluar[1]))
    layak.append(max(hasil[0], keluar[2]))
    layak.append(max(hasil[1], keluar[2]))

    tLayak = max(tidak[0], tidak[1], tidak[2])
    pTimbang = max(tengah[0], tengah[1], tengah[2])
    yLayak = max(layak[0], layak[1], layak[2])

    return [tLayak, pTimbang, yLayak]

#DEFUZZIFIKASI yStar
def defuzzifikasi(inferensi) :
   
    return ((40 * inferensi[0]) + (65 * inferensi[1]) + (80 * inferensi[2]))/(inferensi[0] + inferensi[1] + inferensi[2])

#Mengurutkan nilai yStar yang teratas beserta indexnya
def pantas(yStar,index):
 
	return [i for _, i in sorted(zip(yStar,index), reverse=True)]

#MAIN PROGRAM
readData()

for row in data_mhs :
    luar = pengeluaran(row['pengeluaran'], batasLuarA, batasLuarB, batasLuarC, batasLuarD)
    hasil = penghasilan(row['penghasilan'], batasA, batasB, batasC, batasD)
    inf = inferensi(hasil, luar)
    yStar.append(defuzzifikasi(inf))
    index.append(row['id'])

hasil = pantas(yStar,index)
lyk = hasil[:20]

#Melakukan sorting agar output id tertata dari kecil ke besar
lyk.sort(reverse=False)

#Menampilkan hasil akhir
print('--------------------------------------')
print('| 20 Orang Penerima Potongan BPP 50% |')
print('--------------------------------------')
print("------")
print("| id |")
print("------")
for id in lyk:
	print(id)
    
# MEMBUAT FILE BARU EXCEL DENGAN XLWT
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Bantuan')

row = 1
col = 0

worksheet.write(0, col, 'Id')

for idx in lyk :
    worksheet.write(row, col, int(idx))
    row += 1

print('Files Saved!')
workbook.save('Bantuan.xls')