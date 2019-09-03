import bs4
import xlsxwriter
import datetime
import sys, os

#get file's path
xml_file = sys.argv[1]

#Lay phan mo rong cua file
#if (xml_file[-4:] != ".xml"):
#    print("Ứng dụng không hỗ trợ kết xuất file này")
#    os.system("pause")
#    os.system("exit")
try:
    f1 = open(xml_file, 'r', encoding='utf8') #open file
    data = f1.read() #doc noi dung file vao data
    f1.close()

    soup = bs4.BeautifulSoup(data,'xml')

    str_now = str(datetime.datetime.now())[0:19].replace(':','-') #string: nam-thang-ngay gio-phut-giay
    file_name = soup.find('mst').text + " " + str_now + '.xlsx'
    file_path_tmp = xml_file.rfind('\\') #tim ky tu \ tu phai sang
    file_path = xml_file[0:file_path_tmp]

    #Kiem tra tk gi

    #TK quyet toan TNCN 05 theo TT92
    def qttncn05_tt92(file_path, file_name):
        workbook = xlsxwriter.Workbook(file_name, {'tmpdir': file_path}) #tao file excel cung thu muc voi file xml
        number_format = workbook.add_format({'num_format':'#,##0', 'align':'right'})

        #******** lay du lieu bang ke 05-1 *****
        #Kiểm tra phải có bke 05-1
        if not (soup.find("PLuc_05_1_BK_QTT") is None):
            first_row = 3 #du lieu bat dau ghi tu dong thu 4
            first_col = 0 #du lieu bat dau ghi tu cot 1
            worksheet = workbook.add_worksheet('05-1') #tao sheet bke 05-1
            bke_01 = []
            for x in soup.find_all('PLuc_05_1_BK_QTT'):
                for y in x.find_all('BKeCTietCNhan'):
                    id = y['id'].split('_')[1] #tách chuỗi thành 1 mảng
                    bke_01.append([float(id), y.ct07.text, y.ct08.text, y.ct09.text, y.ct10.text, float(y.ct11.text), float(y.ct12.text), float(y.ct13.text), float(y.ct14.text), float(y.ct15.text), float(y.ct16.text), float(y.ct17.text), float(y.ct18.text), float(y.ct19.text), float(y.ct20.text), float(y.ct21.text), float(y.ct22.text), float(y.ct23.text), float(y.ct24.text)])
            #print(bke_01)

            #for ca_nhan in bke_01:
                #worksheet.write(row, col, ca_nhan[0])
                #worksheet.write(row, col+1, ca_nhan[1])
                #row += 1

            merge_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            worksheet.merge_range('A1:A3', 'STT', merge_format)
            worksheet.merge_range('B1:B3', 'Họ và tên', merge_format)
            worksheet.merge_range('C1:C3', 'Mã số thuế', merge_format)
            worksheet.merge_range('D1:D3', 'Số CMND/Hộ chiếu', merge_format)
            worksheet.merge_range('E1:E3', 'Cá nhân uỷ quyền quyết toán thay', merge_format)
            worksheet.merge_range('F1:H1', 'Thu nhập chịu thuế', merge_format)
            worksheet.merge_range('F2:F3', 'Tổng số', merge_format)
            worksheet.merge_range('G2:H2', 'Trong đó: TNCT làm căn cứ tính giảm thuế', merge_format)
            worksheet.write('G3', 'Làm việc trong KKT', merge_format)
            worksheet.write('H3', 'Theo Hiệp định', merge_format)
            worksheet.merge_range('I1:M1', 'Các khoản giảm trừ', merge_format)
            worksheet.merge_range('I2:I3', 'Số lượng NPT tính giảm trừ', merge_format)
            worksheet.merge_range('J2:J3', 'Tổng số tiền  giảm trừ gia cảnh', merge_format)
            worksheet.merge_range('K2:K3', 'Từ thiện, nhân đạo, khuyến học', merge_format)
            worksheet.merge_range('L2:L3', 'Bảo hiểm được trừ', merge_format)
            worksheet.merge_range('M2:M3', 'Quĩ hưu trí tự nguyện được trừ', merge_format)
            worksheet.merge_range('N1:N3', 'Thu nhập tính thuế', merge_format)
            worksheet.merge_range('O1:O3', 'Số thuế TNCN đã khấu trừ', merge_format)
            worksheet.merge_range('P1:P3', 'Số thuế TNCN được giảm do làm việc trong KKT', merge_format)
            worksheet.merge_range('Q1:S1', 'Chi tiết kết quả quyết toán thay cho cá nhân nộp thuế', merge_format)
            worksheet.merge_range('Q2:Q3', 'Tổng số thuế phải nộp', merge_format)
            worksheet.merge_range('R2:R3', 'Số thuế đã nộp thừa', merge_format)
            worksheet.merge_range('S2:S3', 'Số thuế còn phải nộp', merge_format)

            # add_table(first_row, first_col, last_row, last_col, options)
            #worksheet.add_table('A1:B41', {'data': bke_01})
            worksheet.add_table(first_row, first_col, len(bke_01) + first_row, len(bke_01[0])+first_col-1, {'data': bke_01, 'columns': [{'header':'[06]'},{'header':'[07]'}, {'header':'[08]'}, {'header':'[09]'}, {'header':'[10]','format':number_format}, {'header':'[11]','format':number_format}, {'header':'[12]','format':number_format}, {'header':'[13]','format':number_format}, {'header':'[14]','format':number_format}, {'header':'[15]','format':number_format}, {'header':'[16]','format':number_format}, {'header':'[17]','format':number_format}, {'header':'[18]','format':number_format}, {'header':'[19]','format':number_format}, {'header':'[20]','format':number_format}, {'header':'[21]','format':number_format}, {'header':'[22]','format':number_format}, {'header':'[23]','format':number_format}, {'header':'[24]','format':number_format}]}) 
        #end if
        #***************************************

        #******** lay du lieu bang ke 05-2 *****
        #Kiểm tra phải có bke 05-2
        if not (soup.find("PLuc_05_2_BK_QTT") is None):
            first_row = 3 #du lieu bat dau ghi tu dong thu 4
            first_col = 0 #du lieu bat dau ghi tu cot 1
            worksheet2 = workbook.add_worksheet('05-2') #tao sheet bke 05-2
            bke_02 = [] #[0,'a','b','c',0,0,0,0,0,0,0,0]

            for x in soup.find_all('PLuc_05_2_BK_QTT'):
                for y in x.find_all('BKeCTietCNhan'):
                    id = y['id'].split('_')[1] #tách chuỗi thành 1 mảng
                    bke_02.append([float(id), y.ct07.text, y.ct08.text, y.ct09.text, y.ct10.text, float(y.ct11.text), float(y.ct12.text), float(y.ct13.text), float(y.ct14.text), float(y.ct15.text), float(y.ct16.text), float(y.ct17.text)])

            worksheet2.merge_range('A1:A3', 'STT', merge_format)
            worksheet2.merge_range('B1:B3', 'Họ và tên', merge_format)
            worksheet2.merge_range('C1:C3', 'Mã số thuế', merge_format)
            worksheet2.merge_range('D1:D3', 'Số CMND/Hộ chiếu', merge_format)
            worksheet2.merge_range('E1:E3', 'Cá nhân không cư trú', merge_format)
            worksheet2.merge_range('F1:I1', 'Thu nhập chịu thuế', merge_format)
            worksheet2.merge_range('F2:F3', 'Tổng số', merge_format)
            worksheet2.merge_range('G2:I2', 'Trong đó: TNCT được giảm thuế ', merge_format)
            worksheet2.write('G3', 'Trong đó : TNCT từ phí mua BH nhân thọ, BH không bắt buộc khác của DN BH không thành lập tại Việt Nam cho người lao động', merge_format)
            worksheet2.write('H3', 'Làm việc tại KKT', merge_format)
            worksheet2.write('I3', 'Theo Hiệp định', merge_format)
            worksheet2.merge_range('J1:K1', 'Số thuế thu nhập cá nhân (TNCN) đã khấu trừ', merge_format)
            worksheet2.merge_range('J2:J3', 'Tổng số', merge_format)
            worksheet2.merge_range('K2:K3', 'Trong đó: Số thuế từ phí mua BH nhân thọ, BH không bắt buộc khác của DN BH không thành lập tại Việt Nam cho người lao động', merge_format)
            worksheet2.merge_range('L1:L3', 'Số thuế TNCN được giảm do làm việc tại KKT', merge_format)

            worksheet2.add_table(first_row, first_col, len(bke_02) + first_row, len(bke_02[0])+first_col-1, {'data': bke_02, 'columns': [{'header':'[06]'},{'header':'[07]'}, {'header':'[08]'}, {'header':'[09]'}, {'header':'[10]','format':number_format}, {'header':'[11]','format':number_format}, {'header':'[12]','format':number_format}, {'header':'[13]','format':number_format}, {'header':'[14]','format':number_format}, {'header':'[15]','format':number_format}, {'header':'[16]','format':number_format}, {'header':'[17]','format':number_format}]}) 
        #end if
        #****************************
        #***** close excel file *****
        workbook.close()

        print("Da ket xuat thanh cong file %s" %file_name)
        os.system("pause")
        #******end qttncn05_tt92
    #TK quyet toan TNCN 05 theo TT156
    def qttncn05_tt156(file_path, file_name):
        workbook = xlsxwriter.Workbook(file_name, {'tmpdir': file_path}) #tao file excel cung thu muc voi file xml
        number_format = workbook.add_format({'num_format':'#,##0', 'align':'right'})

        #******** lay du lieu bang ke 05-1 *****
        #Kiểm tra phải có bke 05-1
        if not (soup.find("PLuc_05_1_BK") is None):
            first_row = 3 #du lieu bat dau ghi tu dong thu 4
            first_col = 0 #du lieu bat dau ghi tu cot 1
            worksheet = workbook.add_worksheet('05-1') #tao sheet bke 05-1
            bke_01 = []
            for x in soup.find_all('PLuc_05_1_BK'):
                for y in x.find_all('ChiTietBangKe'):
                    id = y['id'].split('_')[1] #tách chuỗi thành 1 mảng
                    bke_01.append([float(id), y.ct07.text, y.ct08.text, y.ct09.text, y.ct10.text, float(y.ct11.text), float(y.ct12.text), float(y.ct13.text), float(y.ct14.text), float(y.ct15.text), float(y.ct16.text), float(y.ct17.text), float(y.ct18.text), float(y.ct19.text), float(y.ct20.text), float(y.ct21.text), float(y.ct22.text), float(y.ct23.text)])

            merge_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
            worksheet.merge_range('A1:A3', 'STT', merge_format)
            worksheet.merge_range('B1:B3', 'Họ và tên', merge_format)
            worksheet.merge_range('C1:C3', 'Mã số thuế', merge_format)
            worksheet.merge_range('D1:D3', 'Số CMND/Hộ chiếu', merge_format)
            worksheet.merge_range('E1:E3', 'Cá nhân uỷ quyền quyết toán thay', merge_format)
            worksheet.merge_range('F1:H1', 'Thu nhập chịu thuế', merge_format)
            worksheet.merge_range('F2:F3', 'Tổng số', merge_format)
            worksheet.merge_range('G2:H2', 'Trong đó: TNCT làm căn cứ tính giảm thuế', merge_format)
            worksheet.write('G3', 'Làm việc trong KKT', merge_format)
            worksheet.write('H3', 'Theo Hiệp định', merge_format)

            worksheet.merge_range('I1:L1', 'Các khoản giảm trừ', merge_format)

            worksheet.merge_range('I2:I3', 'Tổng số tiền  giảm trừ gia cảnh', merge_format)

            worksheet.merge_range('J2:J3', 'Từ thiện, nhân đạo, khuyến học', merge_format)

            worksheet.merge_range('K2:K3', 'Bảo hiểm được trừ', merge_format)

            worksheet.merge_range('L2:L3', 'Quĩ hưu trí tự nguyện được trừ', merge_format)

            worksheet.merge_range('M1:M3', 'Thu nhập tính thuế', merge_format)

            worksheet.merge_range('N1:N3', 'Số thuế TNCN đã khấu trừ', merge_format)

            worksheet.merge_range('O1:O3', 'Số thuế TNCN được giảm do làm việc trong KKT', merge_format)

            worksheet.merge_range('P1:R1', 'Chi tiết kết quả quyết toán thay cho cá nhân nộp thuế', merge_format)
            worksheet.merge_range('P2:P3', 'Tổng số thuế phải nộp', merge_format)
            worksheet.merge_range('Q2:Q3', 'Số thuế đã nộp thừa', merge_format)
            worksheet.merge_range('R2:R3', 'Số thuế còn phải nộp', merge_format)

            # add_table(first_row, first_col, last_row, last_col, options)
            #worksheet.add_table('A1:B41', {'data': bke_01})
            worksheet.add_table(first_row, first_col, len(bke_01) + first_row, len(bke_01[0])+first_col-1, {'data': bke_01, 'columns': [{'header':'[06]'},{'header':'[07]'}, {'header':'[08]'}, {'header':'[09]'}, {'header':'[10]','format':number_format}, {'header':'[11]','format':number_format}, {'header':'[12]','format':number_format}, {'header':'[13]','format':number_format}, {'header':'[14]','format':number_format}, {'header':'[15]','format':number_format}, {'header':'[16]','format':number_format}, {'header':'[17]','format':number_format}, {'header':'[18]','format':number_format}, {'header':'[19]','format':number_format}, {'header':'[20]','format':number_format}, {'header':'[21]','format':number_format}, {'header':'[22]','format':number_format}, {'header':'[23]','format':number_format}]}) 
        #end if bke 05-1
        #***************************************

        #******** lay du lieu bang ke 05-2 *****
        #Kiểm tra phải có bke 05-2
        if not (soup.find("PLuc_05_2_BK") is None):
            first_row = 3 #du lieu bat dau ghi tu dong thu 4
            first_col = 0 #du lieu bat dau ghi tu cot 1
            worksheet2 = workbook.add_worksheet('05-2') #tao sheet bke 05-2
            bke_02 = []

            for x in soup.find_all('PLuc_05_2_BK'):
                for y in x.find_all('chiTietBangKe'):
                    id = y['id'].split('_')[1] #tách chuỗi thành 1 mảng
                    bke_02.append([float(id), y.ct07.text, y.ct08.text, y.ct09.text, y.ct10.text, float(y.ct11.text), float(y.ct12.text), float(y.ct13.text), float(y.ct14.text), float(y.ct15.text)])

            worksheet2.merge_range('A1:A3', 'STT', merge_format)
            worksheet2.merge_range('B1:B3', 'Họ và tên', merge_format)
            worksheet2.merge_range('C1:C3', 'Mã số thuế', merge_format)
            worksheet2.merge_range('D1:D3', 'Số CMND/Hộ chiếu', merge_format)
            worksheet2.merge_range('E1:E3', 'Cá nhân không cư trú', merge_format)
            worksheet2.merge_range('F1:H1', 'Thu nhập chịu thuế (TNCT)', merge_format)

            worksheet2.merge_range('F2:F3', 'Tổng số', merge_format)

            worksheet2.merge_range('G2:H2', 'Trong đó: TNCT được giảm thuế', merge_format)

            worksheet2.write('G3', 'Làm việc tại KKT', merge_format)

            worksheet2.write('H3', 'Theo Hiệp định', merge_format)

            worksheet2.merge_range('I1:I3', 'Số thuế thu nhập cá nhân (TNCN) đã khấu trừ', merge_format)

            worksheet2.merge_range('J1:J3', 'Số thuế TNCN được giảm do làm việc tại KKT', merge_format)

            worksheet2.add_table(first_row, first_col, len(bke_02) + first_row, len(bke_02[0])+first_col-1, {'data': bke_02, 'columns': [{'header':'[06]'},{'header':'[07]'}, {'header':'[08]'}, {'header':'[09]'}, {'header':'[10]','format':number_format}, {'header':'[11]','format':number_format}, {'header':'[12]','format':number_format}, {'header':'[13]','format':number_format}, {'header':'[14]','format':number_format}, {'header':'[15]','format':number_format}]}) 
        #end if bke 05-2
        #****************************
        #******** lay du lieu bang ke 05-3 *****
        #Kiểm tra phải có bke 05-3


        #end if bke 05-3
        #****************************
        #***** close excel file *****
        workbook.close()

        print("Da ket xuat thanh cong file %s" %file_name)
        os.system("pause")
        #******end qttncn05_tt156

    #maTKhai: 395 - qtoan tncn 05 theo TT92
    if (soup.find('maTKhai').text == '395'):
        qttncn05_tt92(file_path, file_name)
    elif (soup.find('maTKhai').text == '42'):
        qttncn05_tt156(file_path, file_name)
    else:
        print("Ứng dụng chưa hỗ trợ kết xuất tờ khai này")
        os.system("pause")
except:
    print("Lỗi khi kết xuất. Có thể định dạng ứng dụng chưa hỗ trợ. Nhấn phím bất kỳ để thoát!")
    os.system("pause")