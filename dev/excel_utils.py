import pylightxl as xl

def read_excel(file_path):
    db = xl.readxl(fn=file_path)
    ws_name = db.ws_names[0]
    ws = db.ws(ws_name)
    data = list(ws.rows)
    headers = data[0]
    data = data[1:]
    return headers, data

def write_excel(file_path, headers, data):
    db = xl.Database()
    ws_data = [headers] + [[row.get(header, '') for header in headers] for row in data]
    db.add_ws(ws='Sheet1', data=ws_data)
    xl.writexl(db=db, fn=file_path)
