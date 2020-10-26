import csv, traceback, openpyxl, re, glob, datetime, calendar, os
from imbox import Imbox
from common.util import HOST, USERNAME, PASSWORD, OS_Checker

OS = OS_Checker()

if OS == "win":
    download_folder = "C:\\Users\\sariz\\OneDrive\\Desktop\\myPythonScripts\\budgetCSV"
else:
    download_folder = "/Users/sariz/OneDrive/Desktop/myPythonScripts/budgetCSV"  # As example the OS is darwin/linux

if not os.path.isdir("C:\\Users\\sariz\\OneDrive\\Desktop\\myPythonScripts\\budgetCSV"):
    os.makedirs(
        "C:\\Users\\sariz\\OneDrive\\Desktop\\myPythonScripts\\budgetCSV", exist_ok=True
    )

mail = Imbox(
    HOST,
    USERNAME=USERNAME,
    PASSWORD=PASSWORD,
    ssl=True,
    ssl_context=None,
    starttls=False,
)
messages = mail.messages(subject="Transactions", unread=True)

for (uid, message) in messages:
    mail.mark_seen(uid)

    for idx, attachment in enumerate(message.attachments):
        try:
            att_fn = attachment.get("filename")
            download_path = f"{download_folder}/{att_fn}"
            print(download_path)
            with open(download_path, "wb") as fp:
                fp.write(attachment.get("content").read())

        except Exception as e:
            print(e)

mail.logout


# rename the downloaded file to Transactions.csv
for fileName in glob.glob(
    "C:\\users\\sariz\\onedrive\\desktop\\myPythonScripts\\budgetCSV\**\*.csv",
    recursive=True,
):

    dst = "C:\\users\\sariz\\onedrive\\desktop\\myPythonScripts\\budgetCSV\\Transactions.csv"

    os.replace(fileName, dst)  ###

    # convert the csv file into an xlsx
    wb = openpyxl.Workbook()
    ws = wb.active

    os.chdir("C:\\users\\sariz\\onedrive\\desktop\\myPythonScripts\\budgetCSV")

    with open("Transactions.csv") as f:
        reader = csv.reader(f, delimiter=":")
        for row in reader:
            ws.append(row)

    os.remove(
        "C:\\users\\sariz\\onedrive\\desktop\\myPythonScripts\\budgetCSV\\Transactions.csv"
    )

    wb.save("workingSheet.xlsx")

    wb = openpyxl.load_workbook("WorkingSheet.xlsx")

    date = []
    category = []
    amount = []
    remarks = []

    # looks through entries and assign them to corresponding variable lists
    for i in range(2, 100):
        item1 = ws.cell(row=i, column=1).value
        try:
            entries = re.split("\t", item1)

            date.append(entries[0])
            category.append(entries[1])
            amount.append(entries[2])
            remarks.append(entries[3])

            iterations = int(i)

        except TypeError:
            pass
            break

    wb.close()

    # finding out the date and year of the entries
    dateNow = date[0]
    dtArgs = dateNow.split(r"/")
    intDT = [int(i) for i in dtArgs]

    today = datetime.date(intDT[2], intDT[1], intDT[0])

    sheetName = calendar.month_name[today.month] + " " + str(today.year)

    # start budget sheet entries
    os.chdir("c:\\users\\sariz\\onedrive\\desktop")
    wb = openpyxl.load_workbook("budget sheet.xlsx")

    if sheetName not in wb.sheetnames:
        wb.create_sheet(sheetName)

    ws1 = wb[sheetName]
    ws1["B2"] = "Date"
    ws1["C2"] = "Category"
    ws1["D2"] = "Amount"
    ws1["E2"] = "Remarks"

    itemLists = [date, category, amount, remarks]

    rowNum = 3
    columnNum = 2

    for itemNum in range(len(itemLists)):
        for itemScroll in range(len(date)):
            ws1.cell(row=rowNum, column=columnNum).value = itemLists[itemNum][
                itemScroll
            ]
            rowNum += 1
        columnNum += 1
        rowNum = 3

    wb.save("budget sheet.xlsx")
