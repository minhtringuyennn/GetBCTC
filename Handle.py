import requests, json, math
import pandas as pd

# Load data from API
class FetchData():
    def __init__(self, symbol, type, year, checked, count):
        self.symbol = symbol
        
        if type == 'cân đối kế toán':
            self.type = 1
        if type == 'kết quả kinh doanh':
            self.type = 2
        if type == 'lưu chuyển tiền tệ (trực tiếp)':
            self.type = 3
        if type == 'lưu chuyển tiền tệ (gián tiếp)':
            self.type = 4

        year = str(year)
        self.year = year[:-6]
        self.quarter = math.ceil(int(year[5:-3])/3)
        
        if checked == True:
            self.quarter = 0
            
        self.count = count
        
        print(self.symbol, self.type, self.year, self.quarter, self.count)

    def fetchBCTC(self):
        # API
        FIREANT_API = 'https://www.fireant.vn/api/Data/Finance/LastestFinancialReports'
        
        params = {
            "symbol": self.symbol,
            "type": self.type,
            "year": self.year,
            "quarter": self.quarter,
            "count": self.count
        }
        
        # Get reponse from API
        request = requests.get(FIREANT_API, params=params).json()

        try:
            # Filter unnecessary elements from API
            for element in request:
                element.pop('ID', None)
                element.pop('ParentID', None)
                element.pop('Expanded', None)
                element.pop('Level', None)
                element.pop('Field', None)
                for e in element['Values']:
                    period = e['Period']
                    element.update({period: int(e['Value'] or 0)})
                element.pop('Values', None)
        except:
            return None

        # JSON reponse --> Pandas dataframe
        df_json = pd.DataFrame.from_dict(request)
        # df_json.to_excel('temp_data.xlsx')
        
        return df_json