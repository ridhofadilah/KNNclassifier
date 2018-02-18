import xlrd
import xlsxwriter

def euclideanDistance(titikA,titikB):
	tempt = 0
	for i in range (1,len(titikA)-1):
		tempt += ((titikA[i]-titikB[i])**2)
	hasil = [(tempt)**0.5,titikB[len(titikB)-1]]
	return hasil

def kesimpulan(nilai,k):
	yes, no = 0,0
	for i in range(0,k):
		if ((nilai[i])[1]==1):
			yes += 1
		else:
			no += 1
	if yes > no:
		return 1
	else :
		return 0

def crossValidation(first_sheet,fold,k):
	akurasi = 0
	bagian = (first_sheet.nrows-1) // fold
	for h in range(1,fold+1):
		total = 0
		if (h==1):
			batas1 = (bagian*h)+1
			for i in range(1,batas1):
				nilai = []
				titikA = first_sheet.row_values(i)
				for j in range(batas1,first_sheet.nrows):
					titikB = first_sheet.row_values(j)
					nilai.append(euclideanDistance(titikA,titikB))
				nilai.sort()
				if (kesimpulan(nilai,k) == titikA[len(titikA)-1]):
					total += 1
		elif (h==fold):
			batas1 = (bagian*(h-1))+1	
			for i in range(batas1,first_sheet.nrows):
				nilai = []
				titikA = first_sheet.row_values(i)
				for j in range (1,batas1):
					titikB = first_sheet.row_values(j)
					nilai.append(euclideanDistance(titikA,titikB))
				nilai.sort()
				if (kesimpulan(nilai,k) == titikA[len(titikA)-1]):
					total += 1
		else:
			batas1 = (bagian*(h-1))+1 
			batas2 = (bagian*h)+1
			for i in range(batas1,batas2):
				nilai = []
				titikA = first_sheet.row_values(i)
				for j in range(1,batas1):
					titikB = first_sheet.row_values(j)
					nilai.append(euclideanDistance(titikA,titikB))
				for j in range(batas2,first_sheet.nrows):
					titikB = first_sheet.row_values(j)
					nilai.append(euclideanDistance(titikA,titikB))
				nilai.sort()
				if (kesimpulan(nilai,k) == titikA[len(titikA)-1]):
					total += 1
		akurasi += (total/bagian)*100
	return akurasi/fold	

def KNN(first_sheet,second_sheet,k):
	book = xlsxwriter.Workbook("Hasil.xlsx")
	worksheet = book.add_worksheet()
	for i in range(1,second_sheet.nrows):
		nilai = []
		titikA = second_sheet.row_values(i)
		for j in range(1,first_sheet.nrows):
			titikB = first_sheet.row_values(j)
			nilai.append(euclideanDistance(titikA,titikB))
		nilai.sort()
		if (kesimpulan(nilai,k) == 1):
			worksheet.write(i,0,1)
		else:
			worksheet.write(i,0,0)

def open_file(path):
	book = xlrd.open_workbook(path)
	first_sheet = book.sheet_by_index(0)
	maksAkurasi = 0
	optK = 0
	for i in range (1,100,2):
		akurasi = crossValidation(first_sheet,4,i)
		print (i, akurasi)
		if (akurasi > maksAkurasi):
			maksAkurasi,optK = akurasi,i
	print ("=================")
	print (optK, maksAkurasi)
	second_sheet = book.sheet_by_index(1)
	KNN(first_sheet,second_sheet,optK)	

if __name__ == "__main__":
	path = "Dataset Tugas 3 AI 1718.xlsx"
	open_file(path)