# 파일명: excel_data_processor.py

import openpyxl
from pathlib import Path

class DataProcessor:
    def __init__(self, excel_file_path):
        self.file_path = Path(excel_file_path)
        self.data = None
        self.processed_data = None

    def load_data(self):
        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            ws = wb['Sheet1']
            ranges = ['C4:L4', 'C8:K8', 'C11:H11']
            self.data = []
            for cell_range in ranges:
                data_list = []
                for row in ws[cell_range]:
                    data_list.extend([cell.value for cell in row])
                self.data.append(data_list)
            wb.close()
        except FileNotFoundError:
            print(f"Error: File not found at {self.file_path}")
        except openpyxl.utils.exceptions.InvalidFileException:
            print(f"Error: Invalid Excel file at {self.file_path}")
        except KeyError:
            print("Error: 'Sheet1' not found in the workbook.")

    def process_data(self):
        if self.data is None:
            print("No data loaded. Please call load_data() first.")
            return

        self.processed_data = []
        for i, data_list in enumerate(self.data):
#            print(f"Processing data list {i}: {data_list}")
#            processed = [value * 2 for value in data_list if value is not None]  # 예시: 각 값을 2배로
            self.processed_data.append(data_list)

    def load_and_process_data(self):
        self.load_data()
        self.process_data()

    def get_processed_data(self):
        if self.processed_data is None:
            print("Data has not been processed. Please call load_and_process_data() first.")
            return None
        return self.processed_data

class MainApplication:
    def __init__(self, file_path):
        self.processor = DataProcessor(file_path)

    def run(self):
        self.processor.load_and_process_data()
        
        processed_data = self.processor.get_processed_data()
        if processed_data:
            for i, data_list in enumerate(processed_data):
                print(f"Processed data list {i}: {data_list}")
            
            # 여기에서 처리된 데이터를 사용하여 추가 작업을 수행할 수 있습니다.

# 메인 실행 부분
if __name__ == "__main__":
    file_path = "D:\Python\RC_Section_Calculation/Calc_As_input.xlsx"
    app = MainApplication(file_path)
    app.run()