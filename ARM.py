


import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import Qt, QDate
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QKeySequence,QPixmap, QColor
from PyQt5.QtWidgets import QLabel, QApplication, QWidget, QVBoxLayout
import pandas as pd
import shutil
from func.CL파일읽고정보저장 import *
from func.정보대로워드문서작성 import *
from datetime import datetime, timedelta
import time
from get_recent_file_name import get_latest_matching_file

form_class = uic.loadUiType(f'./ARM_UI.ui')[0]
destination_folder = fr'd:\파이썬결과물저장소\ARM\{datetime.now().strftime("%y%m%d")}'
if not os.path.isdir(destination_folder) :
    os.mkdir(destination_folder)
source_folder = os.path.join(os.path.dirname( os.path.abspath( __file__ ) ),'양식')
korean_days = ["일", "월", "화", "수", "목", "금", "토"]


    
class Info():

    def __init__(self) :
        self.date = ""
        self.name = ""
        self.patch_type = ""
        #self.doc_type = ""


class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("AutoReportMaker 0.1")
        self.statusLabel = QLabel(self.statusbar)

        self.setGeometry(1470,28,400,400)
        self.setFixedSize(450,550)
        
        self.dateedit_project.setDate(QDate.currentDate())

        '''기본값입력■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        #소스가 있는 폴더(r2m 쉐어 포인트)
        self.input_departure_dir_path.setText(fr"c:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스(국내)\KR R2M\2023 3분기\230831 업데이트")
        #self.input_departure_dir_path.setText(fr"C:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스(대만)\TW R2M 2023년\2023년 3분기\230822 패치")
        #최종 목적지 폴더(팀 쉐어 포인트)
        self.input_destination_dir_path.setText(fr"C:\Users\mssung\OneDrive - Webzen Inc\R2M\QA\2023년")
        self.input_resultdir.setText(fr'D:\파이썬결과물저장소\ARM\{datetime.now().strftime("%y%m%d")}')
        '''메뉴탭■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        self.menu_howtouse.triggered.connect(lambda : self.파일열기("사용법_ARM.txt"))
        self.menu_patchnote.triggered.connect(lambda : self.파일열기("패치노트_ARM.txt"))
        #self.combox_country.currentTextChanged.connect(self.set_data_path)


        '''폴더세팅부'''
        #self.btn_makedir.clicked.connect(self.set_folder)
        
        
        self.btn_open_0.clicked.connect(self.open_qa_folder)#open_qa_folder
        self.btn_open_1.clicked.connect(self.open_qa_file)#open_qa_folder
        self.btn_make_0.clicked.connect(lambda : self.create_doc('request'))
        self.btn_make_1.clicked.connect(lambda : self.create_doc('kick'))
        self.btn_make_2.clicked.connect(self.create_qa_result)
        '''빌드파일명'''
        self.btn_getbuild_0.clicked.connect(lambda : self.get_filename('Korea'))
        self.btn_getbuild_1.clicked.connect(lambda : self.get_filename('Taiwan'))



    
    def get_filename(self, nation):
        '''
        nation = Korea / Taiwan
        '''
        #nation = Korea
        nation2 = 'KR' if nation == 'Korea' else 'TW'
        directory = fr'D:\Builds\{nation2}'
        file_list = []
        file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'apk'))
        file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'ipa'))
        file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Live_QA', 'apk'))
        file_list.append(get_latest_matching_file(directory, f'signed_gxs_R2MClient{nation}_Live', 'apk'))
        file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Live_QA', 'ipa'))
        #file_list.append(get_latest_matching_file(directory, f'R2MClient{nation}_Alpha', 'ipa'))
        file_list.append('TestFlight 1.x.x(y)')
        #print(file_list)
        #return file_list
        self.plainTextEdit_buildFileNames.clear()
        for file_name in file_list:
            self.plainTextEdit_buildFileNames.insertPlainText(f'{file_name}\n')
    
        return file_list

    # def set_folder(self):
    #     self.makedir()
    #     self.복붙문서적용()

    #     #country = self.combox_country.currentText()
    #     project_name = self.combox_projectname.currentText()
    #     date = pd.to_datetime(self.dateedit_project.text())
    #     patch_type = self.combox_patchtype.currentText()    
    #     #doc_type = self.combox_doctype.currentText()        


    #     source_directory = self.input_departure_dir_path.text()
    #     #destination_path = self.input_destination_dir_path.text()
    #     destination_path = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}/{date.strftime("%y%m%d")} {patch_type}'
    #     if not os.path.isdir(destination_path):                                                           
    #         os.mkdir(destination_path)
    #     #source_directory = os.path.join(departure_path,f'{date.strftime("%y%m%d")} {patch_type}')
    #     #share_point_path = os.path.join(destination_path,f'{date.strftime("%y%m%d")} {patch_type}')
    #     self.폴더내용물전체복사(source_directory,destination_path)

    #     '''QA_결과보고 문서 작성'''
    #     checklist_file_name = fr'CL_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.xlsx'
    #     checklist_file_path = os.path.join(destination_path, checklist_file_name)
    #     save_qa_info_to_csv(checklist_file_path)

    #     result_file_name = fr'QA_결과 보고_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA 결과.docx'
    #     result_file_path = os.path.join(destination_path, result_file_name)
    #     update_qa_report(result_file_path)
      

    def get_info(self):
        date = pd.to_datetime(self.dateedit_project.text())
        name = self.combox_projectname.currentText()
        patch_type = self.combox_patchtype.currentText()    
        #doc_type = self.combox_doctype.currentText()    

        a = Info()
        a.date = date
        a.name = name
        a.patch_type = patch_type
        #a.doc_type = doc_type

        return a



    def create_qa_result(self):
        project = self.get_info()
        
        target_file_name = fr'QA_결과 보고 문서.docx'
        target_file_name = fr'결과 보고.docx'
        temp_nation = '국내' if 'KR' in project.name else '대만'
        source_path = fr'c:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스({temp_nation})'
        target_path = os.path.join(source_folder, target_file_name)
        checklist_file_name = fr'CL_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA.xlsx'
        #checklist_file_path = os.path.join(destination_path, checklist_file_name)
        checklist_file_path = self.get_latest_file_in_directory(source_path,checklist_file_name)
        #os.system('pause')
        try:
            save_qa_info_to_csv(os.path.normpath(checklist_file_path))
        except PermissionError:
            self.print_log(f"파일이 열려있습니다. : {checklist_file_name}")

        destination_path = self.input_resultdir.text()
        result_file_name = fr'QA_결과 보고_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA 결과.docx'
        result_file_path = os.path.join(destination_path, result_file_name)
        try:
            update_doc_by_info(target_path,result_file_path,f'qa_info.csv')
        except PermissionError:
            self.print_log(f"파일이 열려있습니다. : {checklist_file_name}")




        if self.checkBox_1.isChecked() :
            os.startfile(result_file_path)





    def create_doc(self,target : str):
        '''
        target = kick / request
        '''
        project = self.get_info()
        day_of_week = project.date.strftime("%w")

        if target == "kick" :
            source_file_name = fr'QA_Kick-off 문서.docx'
            source_path = os.path.join(source_folder, source_file_name)

            result_file_name = fr'QA_Kick-off_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA.docx'
            destination_path = os.path.join(destination_folder, result_file_name)

            kick_dict = self.read_dict(os.path.join(source_folder, fr'Kick_dict_KR_R2M.csv'))
            kick_dict['NAME'] = f'{project.name.replace("_"," ")} {project.date.strftime("%y%m%d")} {project.patch_type}'
            kick_dict['START_DATE'] = f'{datetime.now().strftime("%Y.%m.%d")}({korean_days[int(day_of_week)]})'#datetime.now()
            kick_dict['END_DATE'] = f'{(project.date-timedelta(days=1)).strftime("%Y.%m.%d")}({korean_days[int(day_of_week)-1]})'
            kick_dict['PATCH_TYPE'] = project.patch_type
            kick_dict['PATCH_DATE'] = f'{project.date.strftime("%Y.%m.%d")}({korean_days[int(day_of_week)]})'
            kick_dict['CLIENT_ALPHA_0'] = self.get_filename('Korea')[0] if 'KR' in project.name else self.get_filename('Taiwan')[0] 
            kick_dict['CLIENT_ALPHA_1'] = self.get_filename('Korea')[1] if 'KR' in project.name else self.get_filename('Taiwan')[1]
            update_doc_by_info(source_path,destination_path,info_dict= kick_dict)



        elif target == "request" :
            source_file_name = fr'빈 문서.docx'
            source_path = os.path.join(source_folder, source_file_name)

            result_file_name = fr'QA_요청_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA.docx'
            destination_path = os.path.join(destination_folder, result_file_name)

            shutil.copy2(source_path, destination_path)






        if self.checkBox_1.isChecked() :
            os.startfile(destination_path)


    def read_dict(self,csv_file_name):
        
        df_temp = pd.read_csv(csv_file_name, index_col='Key')

        dict_temp = {}
        for key, value in df_temp['Value'].items():
            dict_temp[key] = value
        
        del df_temp

        return dict_temp
        #     file_mappings = [
        #         ('QA_Kick-off 문서.docx', f'QA_Kick-off_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx'),
        #         ('QA_요청 문서.docx', f'QA_요청_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx')
        #         # Add more file mappings here
        #     ]
            
        #     for source_file, destination_file in file_mappings:
        #         source_path = os.path.join(source_folder, source_file)
                
        #         # Copy the file with renaming
        #         shutil.copy2(source_path, destination_path)
                
        #         print(f"File '{source_file}' copied and renamed to '{destination_file}'.")
        #     #QA_Kick-off_KR_R2M_230831 업데이트 QA
        #     result_file_path = os.path.join(destination_path, result_file_name)
        
        # checklist_file_name = fr'CL_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA.xlsx'
        # #checklist_file_path = os.path.join(destination_path, checklist_file_name)
        # checklist_file_path = self.get_latest_file_in_directory(fr'c:\Users\mssung\OneDrive - Webzen Inc',checklist_file_name)
        # #os.system('pause')
        # try:
        #     save_qa_info_to_csv(os.path.normpath(checklist_file_path))
        # except PermissionError:
        #     self.print_log(f"파일이 열려있습니다. : {checklist_file_name}")

        # destination_path = self.input_resultdir.text()
        # result_file_name = fr'QA_결과 보고_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA 결과.docx'
        # result_file_path = os.path.join(destination_path, result_file_name)
        # try:
        #     update_qa_report(result_file_path)
        # except PermissionError:
        #     self.print_log(f"파일이 열려있습니다. : {checklist_file_name}")

    def open_qa_file(self):
        project = self.get_info()
        source_path = fr'c:\Users\mssung\OneDrive - Webzen Inc'
        checklist_file_name = fr'CL_{project.name}_{project.date.strftime("%y%m%d")} {project.patch_type} QA.xlsx'
        checklist_file_path = self.get_latest_file_in_directory(source_path,checklist_file_name)
        
        #if checklist_file_path != None:        
        try:
            os.startfile(checklist_file_path)
        except FileNotFoundError:
            self.popUp(f"The folder at path '{checklist_file_path}' does not exist.")
        except Exception as e:
            self.popUp(f"An error occurred: {e}")


    def open_qa_folder(self):
        project = self.get_info()

        temp_nation = '국내' if 'KR' in project.name else '대만'
        source_path = fr'c:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스({temp_nation})'

        folder_name = f'{project.date.strftime("%y%m%d")} {project.patch_type}'
        folder_path = self.find_folders_by_name(source_path, folder_name)

        try:
            os.startfile(folder_path)
        except FileNotFoundError:
            print(f"The folder at path '{folder_path}' does not exist.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def makedir(self):
        '''
        폴더생성
        '''
        #country = self.combox_country.currentText()
        project_name = self.combox_projectname.currentText()
        date = pd.to_datetime(self.dateedit_project.text())#pd.to_datetime(self.dateedit_project.text()).strftime("%y-%m-%d")
        patch_type = self.combox_patchtype.currentText()        
        #doc_type = self.combox_doctype.currentText()        

        #C:\Users\mssung\OneDrive - Webzen Inc\R2M\QA\2023년        

        dir_path = self.input_destination_dir_path.text()
        total_path = f'{dir_path}\{project_name}\{date.strftime("%y%m%d")} {patch_type}'
        
        
        if not os.path.isdir(total_path):                                                           
            os.mkdir(total_path)


        destination_folder = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}'#/{date.strftime("%y%m%d")} {patch_type}'

        if not os.path.isdir(destination_folder):                                                           
            os.mkdir(destination_folder)

    def 복붙문서적용(self):
        '''
        이메일복붙 등의 문서일 경우 디폴트폴더에서 복사/붙여넣기
        킥오프문서/qa요청문서 등...
        '''
        #country = self.combox_country.currentText()
        project_name = self.combox_projectname.currentText()
        date = pd.to_datetime(self.dateedit_project.text())
        patch_type = self.combox_patchtype.currentText()        
        #doc_type = self.combox_doctype.currentText()        
        
        #dir_path = self.input_destination_dir_path.text()
        #total_path = f'{dir_path}\{country} {project_name}\{date.strftime("%y%m%d")} {patch_type}'
        
        source_folder = f'./디폴트양식_{doc_type}' 
        destination_folder = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}/{date.strftime("%y%m%d")} {patch_type}'
        if not os.path.isdir(destination_folder):                                                           
            os.mkdir(destination_folder)
        try:
            # Make sure the destination folder exists
            # if not os.path.exists(destination_folder):
            #     os.makedirs(destination_folder)
            
            # Define the file mappings
            file_mappings = [
                ('QA_Kick-off 문서.docx', f'QA_Kick-off_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx'),
                ('QA_요청 문서.docx', f'QA_요청_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx')
                # Add more file mappings here
            ]
            
            for source_file, destination_file in file_mappings:
                source_path = os.path.join(source_folder, source_file)
                destination_path = os.path.join(destination_folder, destination_file)
                
                # Copy the file with renaming
                shutil.copy2(source_path, destination_path)
                
                print(f"File '{source_file}' copied and renamed to '{destination_file}'.")
        except Exception as e:
            print(f"An error occurred: {e}")


        # departure_path = self.input_departure_dir_path.text()
        # os.startfile(departure_path)
        # checklist_file_name = fr'CL_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.xlsx'
        # checklist_file_path = os.path.join(departure_path,f'{date.strftime("%y%m%d")} {patch_type}' ,checklist_file_name)

        # destination_file_path = os.path.join(dir_path, f'{country} {project_name}' ,checklist_file_name)
        # shutil.copy(checklist_file_path, destination_file_path)

    def 폴더내용물전체복사(self,source_folder, destination_folder):
        try:
            # Make sure the destination folder exists
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            # Get a list of files in the source folder
            file_list = os.listdir(source_folder)
            
            for file_name in file_list:
                source_path = os.path.join(source_folder, file_name)
                destination_path = os.path.join(destination_folder, file_name)
                
                # Copy the file to the destination folder
                shutil.copy(source_path, destination_path)
                
                print(f"File '{file_name}' copied to '{destination_folder}'.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def 파일열기(self,filePath):
        try:
            os.startfile(filePath)
        except : 
            print("파일 없음 : "+filePath)    

    def get_latest_file_in_directory(self, source_path, target_file):

        def find_latest_file(folder):
            latest_file = None
            latest_time = datetime.min

            for root, dirs, files in os.walk(folder):
                if target_file in files:
                    file_path = os.path.join(root, target_file)
                    modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if modified_time > latest_time:
                        latest_file = file_path
                        latest_time = modified_time

            return latest_file

        latest_file_path = find_latest_file(source_path)
        return latest_file_path
        # if latest_file_path:
        #     # 파일 실행 코드 작성
        #     print(f"가장 최신의 파일 실행: {latest_file_path}")
        #     os.startfile(os.path.normpath(latest_file_path))

        # else:
        #     print(f"'{target_file}' 파일을 찾을 수 없습니다.")


    def find_folders_by_name(self, source_path, folder_name):
        #matching_folders = []
        
        for root, dirs, files in os.walk(source_path):
            if folder_name in dirs:
                folder_path = os.path.join(root, folder_name)
                #matching_folders.append(folder_path)
                return folder_path
        
        self.popUp(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")
        return
        

    source_path = fr'C:\Users\mssung\OneDrive - Webzen Inc\R2M_Build\KR'
    #folder_name = 'YourFolderName'  # 검색할 폴더 이름

    # matching_folders = find_folders_by_name(source_path, folder_name)

    # if matching_folders:
    #     for folder in matching_folders:
    #         print(f"폴더 경로: {folder}")
    # else:
    #     print(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")


    def print_log(self, log): # / - \ / - \ / ㅡ ㄷ
        self.progressLabel.setText(log)
        QApplication.processEvents()

    def popUp(self,desText,titleText="error"):
        msg = QtWidgets.QMessageBox()  
        #msg.setGeometry(1520,28,400,2000)
        msg.setText(desText)
        msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)

        x = msg.exec_()

        
    def import_cache_all(self, any_widget = None):
        try:
            # Load CSV file with tab delimiter and utf-16 encoding
            df = pd.read_csv(cache_path, sep='\t', encoding='utf-16', index_col='key')
            if any_widget == None :
                all_widgets = self.findChildren((QLineEdit,  QComboBox, QCheckBox, QPlainTextEdit, QDateEdit))
            else:
                all_widgets = [self.findChild(any_widget[0] ,any_widget[1])]

            for widget in all_widgets:
                object_name = widget.objectName()
                if object_name in df.index:
                    value = str(df.loc[object_name, 'value'])
                    if isinstance(widget, (QLineEdit,QLabel,QPushButton)):
                        widget.setText(value)
                    elif isinstance(widget, QComboBox):
                        # Set selected index based on the value, adjust as needed
                        index = widget.findText(value)
                        if index != -1:
                            widget.setCurrentIndex(index)
                    elif isinstance(widget, QCheckBox):
                        widget.setChecked(value.lower() == 'true')
                    elif isinstance(widget, QPlainTextEdit):
                        widget.setPlainText(value)
                    elif isinstance(widget, QDateEdit):
                        date_format = "yyyy-MM-dd"
                        date = QDate.fromString(value, date_format)
                        widget.setDate(date)

        except Exception as e:
            print(f"Error importing cache: {e}")

    def export_cache_all(self):
        try:
            data = {'key': [], 'value': []}

            all_widgets = self.findChildren((QLineEdit,  QComboBox, QCheckBox, QPlainTextEdit,QDateEdit))

            for widget in all_widgets:
                value = ""
                if isinstance(widget, (QLineEdit,QLabel,QPushButton,QDateEdit)) :
                    value = widget.text()
                elif isinstance(widget, QComboBox):
                    value = widget.currentText()
                elif isinstance(widget, QCheckBox):
                    value = str(widget.isChecked())
                elif isinstance(widget, QPlainTextEdit):
                    value = widget.toPlainText()

                if value != "":
                    key = widget.objectName()
                    data['key'].append(key)
                    data['value'].append(value)

            df = pd.DataFrame(data)
            df.set_index('key', inplace=True)
            df.to_csv(cache_path, sep='\t', encoding='utf-16')
        except Exception as e:
            print(f"Error exporting cache: {e}")

            

    def closeEvent(self,event):
        print("end")

        self.export_cache_all()
if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()