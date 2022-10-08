from openpyxl import load_workbook
from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware


class Entry(BaseModel):
    site: str
    name: str
    status: str
    date: str
    comment: str
    sheet: str



app = FastAPI()
origins = [
    "http://localhost",
    "http://localhost:8080",
    "http://localhost:8081",
    "http://localhost:8082",
    "https://excel.bryceeppler.com",
    "*"
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)




@app.get("/")
async def root():
    return {"message": "Welcome to the API"}


@app.get("/worksheets")
async def worksheets():
    wb = load_workbook('form-submission-book.xlsx')
    sheets = wb.sheetnames
    wb.close()
    return wb.sheetnames


@app.post("/post-to-sheet")
async def post_to_sheet(entry : Entry):
    # load the excel file, form-submission-book.xlsx
    wb = load_workbook('form-submission-book.xlsx')

    # get the sheet name from the entry
    sheet_name = entry.sheet
    ws = wb[sheet_name]

    # get the next available row
    next_row = ws.max_row + 1

    # add the row
    ws[f'A{next_row}'] = entry.site
    ws[f'B{next_row}'] = entry.name
    ws[f'C{next_row}'] = entry.status
    ws[f'D{next_row}'] = entry.date
    ws[f'E{next_row}'] = entry.comment

    # save the excel file
    wb.save('form-submission-book.xlsx')

    # close the excel file
    wb.close()

    return entry
