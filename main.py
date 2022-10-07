from fastapi import FastAPI
from openpyxl import load_workbook


app = FastAPI()


@app.get("/")
async def root():
    return {"message": "Hello World"}


@app.get("/test")
async def test():
    # load the excel file, form-submission-book.xlsx, and add a row
    wb = load_workbook('form-submission-book.xlsx')
    ws = wb.active
    ws['A2'] = 'Hello World'
    wb.save('form-submission-book.xlsx')
    
    return {"message":"Success adding row to excel file"}
