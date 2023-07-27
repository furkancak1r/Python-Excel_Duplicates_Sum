import time
from win32com.client import constants
import win32com.client
import os


uzantim = ".xlsm"
uzantix = ".xlsx"
script_dir = os.path.dirname(os.path.abspath(__file__))
print("\Excel'e geri dönebilirsiniz.\n\nP1 hücresinden aşağıdaki kodları seçebilirsiniz:\n\nÜrün Ağacı için:\n\n ias => Ias'a yapıştırılacak mekanik+elektrik ürün ağacı \n mekanik => Mekanik ürün ağacı\n elektrik => Elektrik ürün ağacı \n\nMaliyet Çalışması için:\n\n toplam maliyet => Mekanik+elektrik toplam maliyet\n mekanik maliyet => Mekanik maliyet\n elektrik maliyet => Elektrik maliyet \n\nProgram size malzeme koduna göre toplanmış satırları verecektir.")

while True:

    def get_second_matching_cell():
        matching_cells = []
        sheet_name = ""
        match_found = False
        excel = None
        p1Values = ["mekanik maliyet", "elektrik maliyet", "toplam maliyet", "mekanik", "elektrik", "ias"]
        matching_excels = []

        try:
            def get_matching_excels():
                excelDispatch = win32com.client.Dispatch("Excel.Application")
                workbooks = excelDispatch.Workbooks

                for i in range(1, workbooks.Count + 1):
                    workbook = workbooks(i)
                    if workbook.Name:
                        sheets = [sheet.Name for sheet in workbook.Sheets]
                        if "İcmal" in sheets and "KUR" in sheets:
                            matching_excels.append(workbook.Name)

                return matching_excels

            matching_excels_array = get_matching_excels()

            for file_name in matching_excels_array:
                excel = win32com.client.GetObject(Class="Excel.Application")
                wb = excel.Workbooks(file_name)

                sheets = [sheet.Name for sheet in wb.Sheets]
                for sheet in sheets:
                    current_sheet = wb.Sheets[sheet]
                    if current_sheet.Range("P1").Value in p1Values:
                        sheet_name = sheet
                        matching_cell = current_sheet.Range("D1:D1000").Find(What="MLZM. KOD")
                        while matching_cell is not None:
                            if matching_cell.Address not in matching_cells:
                                matching_cells.append(matching_cell.Address)
                                matching_cell = current_sheet.Range("D1:D1000").FindNext(After=matching_cell)
                            else:
                                break
                        if len(matching_cells) >= 2:  # Check if we found at least two matches
                            match_found = True
                            break
                        else:
                            matching_cells = []  # Reset matching_cells list to search for the second match

                    if match_found:
                        break

        except Exception as e:
            return [], "", False, ""

        return matching_cells, sheet_name, match_found, file_name   
    matching_cells, sheet_name, match_found, file_name = get_second_matching_cell()
    if match_found == True:
        sorted_result = sorted(
            matching_cells, key=lambda x: int(x[3:].split("$")[0]))
        # Get a reference to an open Excel application
        excel = win32com.client.GetObject(Class="Excel.Application")
        # Get a reference to the workbook you're interested in
        wb = excel.Workbooks(file_name)
        uzantisiz_file_name = file_name[:-5]

        current_sheet = wb.Sheets[sheet_name]

        # new_wb_ürün_agaci = excel.Workbooks(new_file_name_ürün_agaci)

        # if os.path.isfile(new_file_name_maliyet):
        # dosya var, aç
        #   new_wb_maliyet = excel.Workbooks.Open(new_file_name_maliyet)
        # else:
        # dosya yok, oluştur
        #   new_wb_maliyet = excel.Workbooks.Add()
        #  new_wb_maliyet.SaveAs(new_file_name_maliyet)new_file_name_ürün_agaci = uzantisiz_file_name + " Ürün Ağacı.xlsm"
        new_file_name_maliyet = uzantisiz_file_name + " Maliyet Çalışması.xlsx"
        new_file_name_ürün_agaci = uzantisiz_file_name + " Ürün Ağacı.xlsm"
        file_path_ürün_agaci = os.path.join(
            script_dir, new_file_name_ürün_agaci)
        file_path_maliyet = os.path.join(script_dir, new_file_name_maliyet)

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "mekanik":

            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                # 52: xlOpenXMLWorkbookMacroEnabled)
                new_wb_ürün_agaci.SaveAs(
                    Filename=file_path_ürün_agaci, FileFormat=52)
                wb_ürün_agaci = new_wb_ürün_agaci

           
            second_cell_address = sorted_result[1]
            third_cell_address = sorted_result[2]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                if not row[0].startswith("MUHTELIF_GIDER"):
                    printed_data.append([row[0], row[1], row[2], row[3]])

            # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)

            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " MEKANİK"[:31] or sheet.Name == sheet_name + " MEKANİK":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = sheet_name + " MEKANİK"
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    
                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
            try:
                wb_ürün_agaci.Worksheets('Sayfa1').Delete()
            except Exception as e:
                pass
            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "elektrik":
            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                # 52: xlOpenXMLWorkbookMacroEnabled)
                new_wb_ürün_agaci.SaveAs(
                    Filename=file_path_ürün_agaci, FileFormat=52)
                wb_ürün_agaci = new_wb_ürün_agaci

            second_cell_address = sorted_result[2]
            third_cell_address = sorted_result[3]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append([row[0], row[1], row[2], row[3]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " ELEKTRİK":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = sheet_name + " ELEKTRİK"
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(
                            row+1, 1).Value = content[0]
                        new_sheet.Cells(
                            row+1, 2).Value = content[1]
                        new_sheet.Cells(
                            row+1, 3).Value = content[2]
                        new_sheet.Cells(
                            row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    try:
                        wb_ürün_agaci.Worksheets('Sayfa1').Delete()
                    except Exception as e:
                        pass
                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "ias":

            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                new_wb_ürün_agaci.SaveAs(Filename=file_path_ürün_agaci, FileFormat=52)
                wb_ürün_agaci = new_wb_ürün_agaci

            second_cell_address = sorted_result[1]
            third_cell_address = sorted_result[3]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append([row[0], row[1], row[2], row[3]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row == printed_data[0]:
                    # If the row is the header row, add it to the merged data without summing
                    merged_data[row[0]] = row
                elif row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row
            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " IAS":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                   

                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = sheet_name + " IAS"
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(
                            row+1, 1).Value = content[0]
                        new_sheet.Cells(
                            row+1, 2).Value = content[1]
                        new_sheet.Cells(
                            row+1, 3).Value = content[2]
                        new_sheet.Cells(
                            row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    #excel.Application.Run("'" + wb_ürün_agaci.Name + "'!EnableMacros")


                    # Buton ekleyin ve VBA kodunu butona bağlayın
                    left = new_sheet.Range("F2").Left
                    top = new_sheet.Range("F2").Top
                    width = new_sheet.Range("F2").Width
                    height = new_sheet.Range("F2").Height

                    button = new_sheet.Buttons().Add(left, top, width, height)
                    button.Text = "Kopyala"

                    button.OnAction = wb_ürün_agaci.Name + "!ButonaTiklandigindaKopyala.ButonaTiklandigindaKopyala"
                    def check_and_add_macro(wb, module_name, vba_code):
                        # Kontrol edilecek modülü al
                        existing_module = None
                        for module in wb.VBProject.VBComponents:
                            if module.Type == 1 and module.Name == module_name:
                                existing_module = module
                                break
                        
                        if existing_module:
                            # Modül zaten varsa makroyu eklemeyi durdur
                            pass
                        else:
                            # Modül yoksa yeni modül oluştur ve makro kodunu ekleyin
                            module = wb.VBProject.VBComponents.Add(1)  # 1: vbext_ct_StdModule
                            module.Name = module_name
                            module.CodeModule.AddFromString(vba_code)
                    macro_name = "ButonaTiklandigindaKopyala"

                    # VBA kodunu butona atayın
                    vba_code = '''
                    Sub ButonaTiklandigindaKopyala()
                        Dim sourceSheet As Worksheet
                        Dim sourceRange As Range
                        Dim lastRow As Long

                        ' Kaynak sayfayı belirle (Butonun olduğu sayfa)
                        Set sourceSheet = ActiveSheet

                        ' Kaynak veri aralığını belirle (A2 hücresinden D sütununun en son hücresine kadar olan kısım)
                        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "D").End(xlUp).Row
                        Set sourceRange = sourceSheet.Range("A2:D" & lastRow)

                        ' Verileri kopyala
                        sourceRange.Copy
                    End Sub


                    '''
                    check_and_add_macro(wb_ürün_agaci, macro_name, vba_code)                

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    excel.DisplayAlerts = False
                    wb_ürün_agaci.Worksheets(ws).Delete()
                    excel.DisplayAlerts = True  # Uyarıları geri aç

                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "mekanik maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            second_cell_address = sorted_result[1]
            third_cell_address = sorted_result[2]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,  # adet
                                current_sheet.Cells(
                                    row, second_cell.Column + 3).Value,  # birim
                                current_sheet.Cells(
                                    row, second_cell.Column + 5).Value,  # euro-birim
                                current_sheet.Cells(
                                    row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                    merged_data[row[0]][5] += row[5]

                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " MEKANİK M.":
                        sheet_exists = True
                        break

                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(Before=wb.Sheets(1))
                    name = sheet_name + " MEKANİK M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Mekanik Maliyet:"
                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))
                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    # Set the Euro unit in a specific column

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "elektrik maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            second_cell_address = sorted_result[2]
            third_cell_address = sorted_result[3]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 1).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 2).Value,  # adet
                                 current_sheet.Cells(
                                     row, second_cell.Column + 3).Value,  # birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 5).Value,  # euro-birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                    merged_data[row[0]][5] += row[5]

                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " ELEKTRİK M.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(
                        Before=wb_maliyet.Sheets(1))
                    name = sheet_name + " ELEKTRİK M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                   # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Elektrik Maliyet:"
                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))
                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and current_sheet.Range("M1").Value == "Çarpan" and current_sheet.Range("P1").Value == "toplam maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            second_cell_address = sorted_result[1]
            third_cell_address = sorted_result[3]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 1).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 2).Value,  # adet
                                 current_sheet.Cells(
                                     row, second_cell.Column + 3).Value,  # birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 5).Value,  # euro-birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row == printed_data[0]:
                    # If the row is the header row, add it to the merged data without summing
                    merged_data[row[0]] = row
                elif row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    if isinstance(row[2], (int, float)) and isinstance(row[5], (int, float)):
                        merged_data[row[0]][2] += row[2]
                        merged_data[row[0]][5] += row[5]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " TOPLAM M.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(
                        Before=wb_maliyet.Sheets(1))
                    name = sheet_name + " TOPLAM M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name
                   # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Maliyet:"
                    new_sheet.Cells(2, 9).Value = "Kar Oranı:"
                    new_sheet.Cells(3, 9).Value = "Toplam Çözüm Satış Fiyatı:"

                    karorani = current_sheet.Cells(1, 14).Value
                    cozumsatisfiyati = current_sheet.Cells(1, 15).Value
                    new_sheet.Cells(2, 10).Value = karorani
                    new_sheet.Cells(3, 10).Value = cozumsatisfiyati

                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))

                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"
                    new_sheet.Cells(
                        2, 10).NumberFormat = "0%"
                    new_sheet.Cells(
                        3, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"
                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # -4108 #constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "mekanik":

            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                new_wb_ürün_agaci.SaveAs(Filename=file_path_ürün_agaci)
                wb_ürün_agaci = new_wb_ürün_agaci

            second_cell_address = sorted_result[0]
            third_cell_address = sorted_result[1]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append([row[0], row[1], row[2], row[3]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " MEKANİK M.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = sheet_name + " MEKANİK M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "elektrik":
            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                new_wb_ürün_agaci.SaveAs(Filename=file_path_ürün_agaci)
                wb_ürün_agaci = new_wb_ürün_agaci

            third_cell_row = current_sheet.UsedRange.Rows.Count+1
            second_cell_address = sorted_result[1]
            second_cell_row = int(second_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append([row[0], row[1], row[2], row[3]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " ELEKTRİK.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = sheet_name + " ELEKTRİK."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(
                            row+1, 1).Value = content[0]
                        new_sheet.Cells(
                            row+1, 2).Value = content[1]
                        new_sheet.Cells(
                            row+1, 3).Value = content[2]
                        new_sheet.Cells(
                            row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "ias":

            if os.path.exists(file_path_ürün_agaci):
                # Check if the workbook is already open
                if file_path_ürün_agaci in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_ürün_agaci to the already open workbook
                    wb_ürün_agaci = excel.Workbooks(
                        os.path.basename(file_path_ürün_agaci))
                else:
                    # If the workbook is not already open, open it
                    wb_ürün_agaci = excel.Workbooks.Open(file_path_ürün_agaci)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_ürün_agaci = excel.Workbooks.Add()
                new_wb_ürün_agaci.SaveAs(Filename=file_path_ürün_agaci)
                wb_ürün_agaci = new_wb_ürün_agaci

            third_cell_row = current_sheet.UsedRange.Rows.Count+1
            second_cell_address = sorted_result[0]
            second_cell_row = int(second_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,
                                current_sheet.Cells(row, second_cell.Column + 3).Value])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append([row[0], row[1], row[2], row[3]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row == printed_data[0]:
                    # If the row is the header row, add it to the merged data without summing
                    merged_data[row[0]] = row
                elif row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row
            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_ürün_agaci.Sheets:
                    if sheet.Name == sheet_name + " IAS.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists == False:
                    new_sheet = wb_ürün_agaci.Sheets.Add(
                        Before=wb_ürün_agaci.Sheets(1))
                    name = (sheet_name + " IAS.")
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(
                            row+1, 1).Value = content[0]
                        new_sheet.Cells(
                            row+1, 2).Value = content[1]
                        new_sheet.Cells(
                            row+1, 3).Value = content[2]
                        new_sheet.Cells(
                            row+1, 4).Value = content[3]

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Borders.Weight = 2
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    # Buton ekleyin ve VBA kodunu butona bağlayın
                    left = new_sheet.Range("F2").Left
                    top = new_sheet.Range("F2").Top
                    width = new_sheet.Range("F2").Width
                    height = new_sheet.Range("F2").Height

                    button = new_sheet.Buttons().Add(left, top, width, height)
                    button.Text = "Kopyala"

                    button.OnAction = wb_ürün_agaci.Name + "!ButonaTiklandigindaKopyala.ButonaTiklandigindaKopyala"
                    def check_and_add_macro(wb, module_name, vba_code):
                        # Kontrol edilecek modülü al
                        existing_module = None
                        for module in wb.VBProject.VBComponents:
                            if module.Type == 1 and module.Name == module_name:
                                existing_module = module
                                break
                        
                        if existing_module:
                            # Modül zaten varsa makroyu eklemeyi durdur
                            pass
                        else:
                            # Modül yoksa yeni modül oluştur ve makro kodunu ekleyin
                            module = wb.VBProject.VBComponents.Add(1)  # 1: vbext_ct_StdModule
                            module.Name = module_name
                            module.CodeModule.AddFromString(vba_code)
                    macro_name = "ButonaTiklandigindaKopyala"

                    # VBA kodunu butona atayın
                    vba_code = '''
                    Sub ButonaTiklandigindaKopyala()
                        Dim sourceSheet As Worksheet
                        Dim sourceRange As Range
                        Dim lastRow As Long

                        ' Kaynak sayfayı belirle (Butonun olduğu sayfa)
                        Set sourceSheet = ActiveSheet

                        ' Kaynak veri aralığını belirle (A2 hücresinden D sütununun en son hücresine kadar olan kısım)
                        lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "D").End(xlUp).Row
                        Set sourceRange = sourceSheet.Range("A2:D" & lastRow)

                        ' Verileri kopyala
                        sourceRange.Copy
                    End Sub


                    '''
                    check_and_add_macro(wb_ürün_agaci, macro_name, vba_code)
                else:
                    print(
                        "Karakter sınırlaması 31'i geçiyor!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "mekanik maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            second_cell_address = sorted_result[0]
            third_cell_address = sorted_result[1]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            third_cell_row = int(third_cell_address[3:].split("$")[0])

            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 1).Value,
                                current_sheet.Cells(
                                    row, second_cell.Column + 2).Value,  # adet
                                current_sheet.Cells(
                                    row, second_cell.Column + 3).Value,  # birim
                                current_sheet.Cells(
                                    row, second_cell.Column + 5).Value,  # euro-birim
                                current_sheet.Cells(
                                    row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                    merged_data[row[0]][5] += row[5]

                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " MEKANİK M.":
                        sheet_exists = True
                        break

                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(
                        Before=wb_maliyet.Sheets(1))
                    name = sheet_name + " MEKANİK M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                    # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Mekanik Maliyet:"
                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))
                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    # Set the Euro unit in a specific column

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "elektrik maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            third_cell_row = current_sheet.UsedRange.Rows.Count+1
            second_cell_address = sorted_result[1]
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 1).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 2).Value,  # adet
                                 current_sheet.Cells(
                                     row, second_cell.Column + 3).Value,  # birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 5).Value,  # euro-birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    merged_data[row[0]][2] += row[2]
                    merged_data[row[0]][5] += row[5]

                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " ELEKTRİK M.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(
                        Before=wb_maliyet.Sheets(1))
                    name = sheet_name + " ELEKTRİK M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name

                   # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Elektrik Maliyet:"
                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))
                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:

                    pass

            current_sheet.Range('P1').ClearContents()

        if len(sorted_result) > 1 and (current_sheet.Range("C3").Value == "KORUMA SICAKLIĞI" or current_sheet.Range("C5").Value == "KORUMA SICAKLIĞI") and current_sheet.Range("P1").Value == "toplam maliyet":
            if os.path.exists(file_path_maliyet):
                # Check if the workbook is already open
                if file_path_maliyet in [wb.FullName for wb in excel.Workbooks]:
                    # If the workbook is already open, set wb_maliyet to the already open workbook
                    wb_maliyet = excel.Workbooks(
                        os.path.basename(file_path_maliyet))
                else:
                    # If the workbook is not already open, open it
                    wb_maliyet = excel.Workbooks.Open(file_path_maliyet)
            else:
                # If the file doesn't exist, create a new workbook
                new_wb_maliyet = excel.Workbooks.Add()
                new_wb_maliyet.SaveAs(Filename=file_path_maliyet)
                wb_maliyet = new_wb_maliyet

            second_cell_address = sorted_result[0]
            third_cell_row = current_sheet.UsedRange.Rows.Count+1
            second_cell_row = int(second_cell_address[3:].split("$")[0])
            second_cell = current_sheet.Range(second_cell_address)
            contents = []

            for row in range(second_cell_row, third_cell_row):
                contents.append([current_sheet.Cells(row, second_cell.Column).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 1).Value,
                                 current_sheet.Cells(
                                     row, second_cell.Column + 2).Value,  # adet
                                 current_sheet.Cells(
                                     row, second_cell.Column + 3).Value,  # birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 5).Value,  # euro-birim
                                 current_sheet.Cells(
                                     row, second_cell.Column + 7).Value  # euro-toplam

                                 ])
            contents = [
                content for content in contents if content[0] and content[2]]
            extracted_data = {'Extracted Contents': contents}

            # Initialize an empty list to store the printed data
            printed_data = []

            # Add the extracted data to the printed data list
            for row in extracted_data['Extracted Contents']:
                printed_data.append(
                    [row[0], row[1], row[2], row[3], row[4], row[5]])

                # Merge rows with the same MLZM. KOD value
            merged_data = {}
            for row in printed_data:
                if row == printed_data[0]:
                    # If the row is the header row, add it to the merged data without summing
                    merged_data[row[0]] = row
                elif row[0] in merged_data:
                    # If the row already exists, add the values to the existing row
                    if isinstance(row[2], (int, float)) and isinstance(row[5], (int, float)):
                        merged_data[row[0]][2] += row[2]
                        merged_data[row[0]][5] += row[5]
                else:
                    # If the row is new, add it to the dictionary
                    merged_data[row[0]] = row

            # Delete the old printed data
            del printed_data[:]

            # Convert the merged data back into a list and add it to the printed data
            for row in merged_data.values():
                printed_data.append(row)
            try:
                sheet_exists = False
                for sheet in wb_maliyet.Sheets:
                    if sheet.Name == sheet_name + " TOPLAM M.":
                        sheet_exists = True
                        break

                # Create a new sheet in the workbook
                if sheet_exists != True:
                    new_sheet = wb_maliyet.Sheets.Add(
                        Before=wb_maliyet.Sheets(1))
                    name = sheet_name + " TOPLAM M."
                    if len(name) > 31:
                        print("Sayfa ismi karakter sınırlaması 31'i geçiyor!")
                    else:
                        new_sheet.Name = name
                   # Write the data to the sheet
                    for row, content in enumerate(printed_data):
                        new_sheet.Cells(row+1, 1).Value = content[0]
                        new_sheet.Cells(row+1, 2).Value = content[1]
                        new_sheet.Cells(row+1, 3).Value = content[2]
                        new_sheet.Cells(row+1, 4).Value = content[3]
                        new_sheet.Cells(row+1, 5).Value = content[4]
                        new_sheet.Cells(row+1, 6).Value = content[5]

                    if any(content[4] != 0 for content in printed_data):
                        new_sheet.Columns(
                            5).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    if any(content[5] != 0.0 for content in printed_data):
                        new_sheet.Columns(
                            6).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Cells(1, 5).Value = "Birim Fiyat"
                    new_sheet.Cells(1, 6).Value = "Tutar"
                    new_sheet.Cells(1, 9).Value = "Toplam Maliyet:"
                    total = new_sheet.Application.WorksheetFunction.Sum(
                        new_sheet.Columns(6))
                    new_sheet.Cells(1, 10).Value = total
                    new_sheet.Cells(
                        1, 10).NumberFormat = "#.##0,00 €;-#.##0,00 €"

                    new_sheet.Columns.AutoFit()
                    for cell in new_sheet.UsedRange:
                        if cell.Value:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                        if cell.Column >= 1 and cell.Column <= 6:
                            cell.VerticalAlignment = -4108  # constants.xlCenter
                            cell.Interior.ThemeColor = 8  # constants.xlThemeColorAccent4
                            cell.Interior.TintAndShade = 0.8
                            cell.Borders.Weight = 2
                    new_sheet.Range(
                        "A:B").HorizontalAlignment = -4131  # constants.xlLeft
                    new_sheet.Range(
                        "C:D").HorizontalAlignment = -4108  # constants.xlCenter
                    # new_sheet.Cells.HorizontalAlignment = constants.xlHAlignCenter
                    new_sheet.Cells(
                        1, 1).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 2).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 5).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(
                        1, 6).HorizontalAlignment = -4108  # constants.xlCenter
                    new_sheet.Cells(1, 1).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 2).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 3).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 4).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 5).Interior.TintAndShade = 0.4
                    new_sheet.Cells(1, 6).Interior.TintAndShade = 0.4

                    new_sheet.Cells(1, 1).Font.Bold = True
                    new_sheet.Cells(1, 2).Font.Bold = True
                    new_sheet.Cells(1, 3).Font.Bold = True
                    new_sheet.Cells(1, 4).Font.Bold = True
                    new_sheet.Cells(1, 5).Font.Bold = True
                    new_sheet.Cells(1, 6).Font.Bold = True

                    row_count = new_sheet.UsedRange.Rows.Count
                    for row in range(1, row_count+1):
                        if not new_sheet.Cells(row, 5).Value and not new_sheet.Cells(row, 6).Value:
                            new_sheet.Range(new_sheet.Cells(row, 1), new_sheet.Cells(
                                row, 6)).Interior.ColorIndex = 3
                        if new_sheet.Cells(row, 9).Value:
                            new_sheet.Cells(row, 9).Interior.TintAndShade = 0.4
                            new_sheet.Cells(row, 9).Font.Bold = True

                else:
                    print(
                        "Oluşturulacak sekme zaten var!")
            except TypeError:
                print(
                    "Oluşturulacak sekme zaten var!")
         # çalışma sayfalarının adlarını bir liste olarak tanımla
            worksheets = ['Sayfa1', 'Sheet1']

            # listedeki her çalışma sayfası için silmeyi dene
            for ws in worksheets:
                try:
                    wb_ürün_agaci.Worksheets(ws).Delete()
                except Exception as e:
                    pass

            current_sheet.Range('P1').ClearContents()
    time.sleep(6)
