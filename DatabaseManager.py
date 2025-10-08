import pandas

# RULES:
# 1. JANGAN GANTI NAMA CLASS ATAU FUNGSI YANG ADA
# 2. JANGAN DELETE FUNGSI YANG ADA
# 3. JANGAN DELETE ATAU MENAMBAH PARAMETER PADA CONSTRUCTOR ATAU FUNGSI
# 4. GANTI NAMA PARAMETER DI PERBOLEHKAN
# 5. LARANGAN DI ATAS BOLEH DILANGGAR JIKA ANDA TAU APA YANG ANDA LAKUKAN (WAJIB BISA JELASKAN)
# GOODLUCK :)

class excelManager:
    def __init__(self,filePath:str,sheetName:str="Sheet1"):
        self.__filePath = filePath
        self.__sheetName = sheetName
        self.__data = pandas.read_excel(filePath,sheet_name=sheetName)
            
    
    def insertData(self,newData:dict,saveChange:bool=False):
        # kerjakan disini
        # clue cara insert row: df = pandas.concat([df, pandas.DataFrame([{"NIM":0,"Nama":"Udin","Nilai":1000}])], ignore_index=True)
        # Pastikan newData memiliki struktur yang benar
        required_keys = ["NIM", "Nama", "Nilai"]
        
        # Validasi data input
        for key in required_keys:
            if key not in newData:
                return False
        
        # Cek duplikasi NIM
        existing_data = self.getData("NIM", str(newData["NIM"]))
        if existing_data is not None:
            return False
        
        # Validasi nama tidak mengandung angka
        nama = str(newData["Nama"])
        if any(char.isdigit() for char in nama):
            return False
        
        # Pastikan tipe data konsisten
        formatted_data = {
            "NIM": str(newData["NIM"]).strip(),
            "Nama": str(newData["Nama"]).strip(),
            "Nilai": int(newData["Nilai"])
        }
        
        # Insert data baru dengan struktur yang konsisten
        new_row = pandas.DataFrame([formatted_data])
        self.__data = pandas.concat([self.__data, new_row], ignore_index=True)
        
        # Pastikan struktur kolom tetap 3 kolom
        self.__data = self.__data[["NIM", "Nama", "Nilai"]]
        
        if saveChange: 
            self.saveChange()
        
        return True
    
    def deleteData(self, targetedNim:str,saveChange:bool=False):
        # kerjakan disini
        # clue cara delete row: df.drop(indexBaris, inplace=True); contoh: df.drop(0,inplace=True)
        
         # Cari data berdasarkan NIM yang ditargetkan
        data_to_delete = self.getData("NIM", targetedNim)
        
        # Jika NIM tidak ditemukan, return False
        if data_to_delete is None:
            return False
        
        # Dapatkan index baris yang akan dihapus
        row_index = data_to_delete["Row"]
        
        # Hapus baris dari dataframe
        self.__data.drop(row_index, inplace=True)
        
        # Reset index dataframe agar berurutan kembali
        self.__data.reset_index(drop=True, inplace=True)
        
        # Jika saveChange True, simpan perubahan ke file excel
        if saveChange: 
            self.saveChange()
        
        return True  # Return True jika berhasil dihapus
    
    def editData(self, targetedNim:str, newData:dict,saveChange:bool=False) -> dict:
        # kerjakan disini
        # clue cara ganti value: df.at[indexBaris,namaKolom] = value; contoh: df.at[0,ID] = 1
        # Cari data lama
        old_data = self.getData("NIM", targetedNim)
        if old_data is None:
            return None
        
        row_index = old_data["Row"]
        
        # Validasi NIM baru tidak duplikat (jika diubah)
        if targetedNim != newData["NIM"]:
            existing_data = self.getData("NIM", newData["NIM"])
            if existing_data is not None:
                return None
        
        # Validasi nama tidak mengandung angka
        nama_baru = str(newData["Nama"])
        if any(char.isdigit() for char in nama_baru):
            return None
        
        # Format data baru
        formatted_data = {
            "NIM": str(newData["NIM"]).strip(),
            "Nama": str(newData["Nama"]).strip(),
            "Nilai": int(newData["Nilai"])
        }
        
        # Update data
        for column, value in formatted_data.items():
            self.__data.at[row_index, column] = value
        
        if saveChange: 
            self.saveChange()
        
        return self.getData("NIM", newData["NIM"])
    
                    
    def getData(self, colName:str, data:str) -> dict:
        collumn = self.__data.columns # mendapatkan list dari nama kolom tabel
        
        # cari index dari nama kolom dan menjaganya dari typo atau spasi berlebih
        collumnIndex = [i for i in range(len(collumn)) if (collumn[i].lower().strip() == colName.lower().strip())] 
        
        # validasi jika input kolom tidak ada pada data excel
        if (len(collumnIndex) != 1): return None
        
        # nama kolom yang sudah pasti benar dan ada
        colName = collumn[collumnIndex[0]]

        resultDict = dict() # tempat untuk hasil
        
        for i in self.__data.index: # perulangan ke baris tabel
            cellData = str(self.__data.at[i,colName]) # isi tabel yand dijadikan str
            if (cellData == data): # jika data cell sama dengan data input
                for col in collumn: # perulangan ke nama-nama kolom
                    resultDict.update({str(col):str(self.__data.at[i,col])}) # masukan data {namaKolom : data pada cell} ke resultDict
                resultDict.update({"Row":i}) # tambahkan row nya pada resultDict
                return resultDict # kembalikan resultDict
        
        return None
    
    def saveChange(self):
        self.__data.to_excel(self.__filePath, sheet_name=self.__sheetName , index=False)
    
    def getDataFrame(self):
        return self.__data
