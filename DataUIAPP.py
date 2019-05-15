"""
@author: Rahul Tevatia
"""
from tkinter import *
import tkinter.messagebox
import pyodbc
from datetime import datetime, timedelta
from openpyxl import Workbook

try:
    conn = pyodbc.connect('Driver={SQL Server Native Client 11.0};'
                          'Server={};'    #Add server info
                          'Database={};'
                          'Trusted_Connection=yes;')
    cursor = conn.cursor()
except pyodbc.Error as err:
    tkinter.messagebox.showinfo("Error", "Connection Problem or Access Denied")


class Window:
    def __init__(self, master):
        self.master = master
        self.master.title("Data Search Utility")
        self.master.geometry("700x200+0+0")
        f1 = Frame(master, height=400, width=200, relief=RAISED, border=2)
        f1.grid(row=0, column=0, padx=10, pady=2)
        f2 = Frame(master, height=200, width=200, border=2)
        f2.grid(row=1, column=0, padx=10, pady=2)
#        f3 = Frame(master, height=50, width=400, relief=RAISED, border=2)
#        f3.grid(row=2, column=0, padx=10, pady=2)

        self.var1 = IntVar()
        l1 = Radiobutton(f1, text="Plan No.(T0..)", variable=self.var1, value=1, fg="gray20")
        l1.grid(row=0, column=0, padx=10, pady=2, sticky="w")
        self.var1.set("1")

        l2 = Label(f1, text="Begin Date(eg.YYYY-MM-DD)", fg="gray20")
        l2.grid(row=0, column=2, padx=10, pady=2, sticky="w")

        l3 = Radiobutton(f1, text="Product", variable=self.var1, value=2, fg="gray20")
        l3.grid(row=5, column=0, padx=10, pady=2, sticky="w")

        l4 = Label(f1, text="End Date-Optional(eg.YYYY-MM-DD)", fg="gray20")
        l4.grid(row=5, column=2, padx=10, pady=2, sticky="w")

        self.Plan_text = StringVar()
        self.e1 = Entry(f1, textvariable=self.Plan_text)
        self.e1.grid(row=0, column=1, padx=10, pady=2)

        self.Txnstdate_text = StringVar()
        self.e2 = Entry(f1, textvariable=self.Txnstdate_text)
        self.e2.grid(row=0, column=3, padx=10, pady=2)

        self.Txnendate_text = StringVar()
        self.e4 = Entry(f1, textvariable=self.Txnendate_text)
        self.e4.grid(row=5, column=3, padx=10, pady=2)

        b1 = Button(f2, text="Get Transaction History", width=20, command=self.get_txn_hist, bg="gray70")
        b1.grid(row=0, column=0, padx=10, pady=2)

        b2 = Button(f2, text="Get FundSourceBalance", width=20, command=self.get_fsb, bg="gray70")
        b2.grid(row=0, column=1, padx=10, pady=2)

        b3 = Button(f2, text="Get Client Details", width=20, command=self.get_client_details, bg="gray70")
        b3.grid(row=1, column=0, padx=10, pady=2)

        b4 = Button(f2, text="Get Allocation Details", width=20, command=self.get_alloc_details, bg="gray70")
        b4.grid(row=1, column=1, padx=10, pady=2)

        b5 = Button(f2, text="Clear All", width=12, command=self.cleardata, bg="gray70")
        b5.grid(row=2, column=0, padx=10, pady=2)

        b6 = Button(f2, text="Close", width=12, command=master.destroy, bg="gray70")
        b6.grid(row=2, column=1, padx=10, pady=2)

    def cleardata(self):
        self.e1.delete(0, END)
        self.e2.delete(0, END)
#        self.e3.delete(0, END)
        self.e4.delete(0, END)
        self.var1.set("1")

    def get_txn_hist(self):
        wb = Workbook()
        procflag = False
        selection = self.var1.get()
        datest = dateen = ''
        if (self.Plan_text.get() == '') and (self.Txnstdate_text.get() == '') and (self.Txnendate_text.get() == ''):
            tkinter.messagebox.showinfo("Error", "Please Enter Plan No or Product")
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() == '') and (self.Txnendate_text.get() == ''):
            datest = '1900-01-01'
            dateen = datetime.date(datetime.now())
            procflag = True
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() != '') and (self.Txnendate_text.get() == ''):
            datest = dateen = self.Txnstdate_text.get()
            procflag = True
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() != '') and (self.Txnendate_text.get() != ''):
            datest = self.Txnstdate_text.get()
            dateen = self.Txnendate_text.get()
            procflag = True
        else:
            tkinter.messagebox.showinfo("Error", "Please check Details")

        if procflag:
            if selection == 1:
                partplan = self.Plan_text.get()
                prod = ''
            else:
                partplan = ''
                prod = self.Plan_text.get()

            try:
                ws = wb.active
                ws.title = 'Client_Transaction_history'
#            ws.path()
                column_names = ["Product", "Application_number", "Plan_number", "Customer_name", "Plan issue date",
                                "Transaction Type", "Description", "Source Fund Code", "Trade Date", "Run Date",
                                "Run Time", "Reversal Code", "Sequence Number", "Post Number", "Base Curr Code",
                                "Target Curr Code", "Base to target Rate", "Fund Code", "Strategy Code",
                                "Life Stages Code", "Amount", "Unit Price", "No of Units", "Transaction Description",
                                "Comment", "Omni Activity Code", "Source System Transaction Type ID", "Source Amount",
                                "Exchange Type Code", "Source Contribution Type", "Reference Plan Number",
                                "Request Reference", "Pending Status"]
                ws.append(column_names)
                sql_query = """Select PLN.PRD_CODE, PLN.APL_CODE, PLN.PLN_NUM, 
                       CONCAT(UPF.firstname,' ',UPF.first1,' ', UPF.first2, ' ', UPF.surname),
                       PLN.ISSU_DT, TP.TXN_TP_ID, TP.DSC,
                       FM.SRC_STM_FND_CODE,TRD_DT,RUN_DT,RUN_TM,RVRS_CODE,SEQ_NUM,PST_NUM,BASE_CCY_CODE,TRGT_CCY_CODE,
                       BASE_TO_TRGT_RATE,TXN.FND_CODE,TXN.IVSM_STRTG_CODE,TXN.LFE_STAGES_CODE,AMT,UNIT_PRC,
                       NUM_OF_UNITS,TXN_DSC,TXN.CMNT,OMNI_AVY_CODE,SRC_STM_TXN_TP_ID,SRC_AMT,EXG_TYPE_CODE,
                       SRC_CTB_TP_CODE,REFR_PLN_NUM,REQ_REFR,PNDG_ST
                       from ODS.TXN_HIST TXN JOIN ODS.PLN PLN ON TXN.PLN_NUM = PLN.PLN_NUM
                       JOIN MDM.TXN_TP TP ON TXN.TXN_TP_ID = TP.TXN_TP_ID
                       JOIN MDM.FND_MPNG FM ON PLN.PRD_CODE = FM.PRD_CODE AND FM.FND_CODE = TXN.FND_CODE AND
                       FM.IVSM_STRTG_CODE = TXN.IVSM_STRTG_CODE AND FM.LFE_STAGES_CODE = TXN.LFE_STAGES_CODE
                       JOIN dbo.unplatform UPF ON  PLN.PLN_NUM = UPF.planid AND PLN.APL_CODE = UPF.applicant
                       where (PLN.PLN_NUM = ? OR PLN.PRD_CODE = ?) AND TRD_DT >= ? AND TRD_DT <= ?"""
                input1 = (partplan, prod, datest, dateen)
                cursor.execute(sql_query, input1)
                result = cursor.fetchall()
                if len(result) == 0:   # this not working
                    tkinter.messagebox.showinfo("Error", "Check Input Details! No Data Found")
                else:
                    for row in result:
                        ws.append(tuple(row))

                    workbook_name = self.Plan_text.get() + "_Transaction history"
                    wb.save(workbook_name + ".xlsx")
                    tkinter.messagebox.showinfo("Error", "Transaction File Created and saved on your /Desktop")
            except pyodbc.Error as err1:
                tkinter.messagebox.showinfo("Error", "There is some problem in fetching data!! Check Connection or"
                                                     " Input Details")
                print(err1)

    def get_fsb(self):
        wb = Workbook()
        procflag = False
#        selection = self.var1.get()
        datest = dateen = ''
        if (self.Plan_text.get() == '') and (self.Txnstdate_text.get() == '') and (self.Txnendate_text.get() == ''):
            tkinter.messagebox.showinfo("Error", "Please Enter Plan No")
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() == '') and (self.Txnendate_text.get() == ''):
            if len(self.Plan_text.get()) < 10:
                tkinter.messagebox.showinfo("Error", "Please Check Input Details and Enter Plan No./Date")
            else:
#                datest = '1900-01-01'
                datest = datetime.date(datetime.now() - timedelta(1))
                dateen = datetime.date(datetime.now())
                procflag = True
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() != '') and (self.Txnendate_text.get() == ''):
            if len(self.Plan_text.get()) < 10:
                tkinter.messagebox.showinfo("Error", "Please Check Input Details and Enter Plan No./Date")
            else:
                datest = dateen = self.Txnstdate_text.get()
                procflag = True
        elif (self.Plan_text.get() != '') and (self.Txnstdate_text.get() != '') and (self.Txnendate_text.get() != ''):
            if len(self.Plan_text.get()) < 10:
                tkinter.messagebox.showinfo("Error", "Please Check Input Details and Enter Plan No./Date")
            else:
                datest = self.Txnstdate_text.get()
                dateen = self.Txnendate_text.get()
                procflag = True
        else:
            tkinter.messagebox.showinfo("Error", "Please check Details")

        if procflag:
            try:
                ws = wb.active
                ws.title = 'Client_Fund_Source_Balance'
                column_names = ["Product", "Plan Name", "Plan No", "Omni Participant Id", "Client Name",
                                "Source Name", "ISIN", "Strategy Code", "Investment Code", "No of Units",
                                "Fund Curr", "Unit Price", "Price NAV Date", "Fund Value", "Un-invested Cash",
                                "Pending Credits", "Pending Debits", "Advisor", "Load Date"]
                sql_query = """Select * from dbo.fundsrcbal
                               where partplanid = ?  AND loaddate >= ? AND loaddate <= ?"""
                input1 = (self.Plan_text.get(), datest, dateen)
                cursor.execute(sql_query, input1)
                result = cursor.fetchall()
                #            ws.path()
                if len(result) == 0:  # this not working
                    tkinter.messagebox.showinfo("Error", "Check Input Details! No Data Found")
                else:
                    ws.append(column_names)
                    for row in result:
                        ws.append(tuple(row))

                    workbook_name = self.Plan_text.get() + "_FundSource"
                    wb.save(workbook_name + ".xlsx")
                    tkinter.messagebox.showinfo("Error", "FSB File Created and saved on your /Desktop")
            except pyodbc.Error as err1:
                tkinter.messagebox.showinfo("Error", "There is some problem in fetching data!! Check Connection or "
                                                     "Input Details")
                print(err1)

    def get_alloc_details(self):
        tkinter.messagebox.showinfo("Error", "Sorry!!This functionality is under development")
#        if self.Plan_text.get() == '':
#            tkinter.messagebox.showinfo("Error", "Please Enter Plan No(T0..)")
#        else:
#            try:
#                column_names = ["Source", "Contribution Type", "Strategy", "Allocation Percentage"]
#                cursor.execute("""Select distinct source, contribtype, tiscode, cast(sum(fixalloc)*100 as varchar(10))
#                               from dbo.unplatformalloc where partPlanId = ?
#                               group by source, contribtype, tiscode""", self.Plan_text.get())
#                sql_data = cursor.fetchall()
#                if len(sql_data) == 0:
#                    tkinter.messagebox.showinfo("Error", "Please check Plan No and try again")
#                else:
#                    root2 = Toplevel(self.master)
#                    root2.geometry('1150x350+0+0')
#                    root2.title('Client Allocation Details')
#                    frm1 = Frame(root2, width=800, height=150, relief=RAISED, border=5)
#                    frm1.grid(row=0)
#                    Label(frm1, text="Strategy Details", font=("Times New Roman", 18), fg='gray40', anchor=W).grid(
#                                                                                                            row=0,
#                                                                                                            column=0,
#                                                                                                            sticky="w")
#                    i = 0
#                    for row in column_names:
#                        Label(frm1, text=row, anchor=W, font=("Times New Roman", 12), fg='royal blue').grid(row=i + 1,
#                                                                                                            column=0,
#                                                                                                            sticky="w")
#                        i += 1
#                    j = 0
#                    for row in sql_data:
#                        i = 0
#                        print(row)
#                        for data in row:
#                            Label(frm1, text=data, anchor=W, font=("Times New Roman", 12)).grid(row=i + 1,
#                                                                                                column=j + 1,
#                                                                                                sticky="w")
#                            i += 1
#                        j += 1
#
#                    Label(frm1, text="Fund Details", anchor=W,
#                          font=("Times New Roman", 18), fg='gray40').grid(row=0, column=4, sticky="w")
#                    column_names = ["Investment", "Investment Name", "Omni Fund Code", "Allocation Percentage"]
#                    i = 0
#                    for row in column_names:
#                        Label(frm1, text=row, anchor=W, font=("Times New Roman", 12), fg='royal blue').grid(row=i + 1,
#                                                                                                            column=4,
#                                                                                                            sticky="w")
#                        i += 1
#
#                    cursor.execute("""Select distinct fundcode, recordtype, CONCAT(InvCode, SourceCode),
#                                   cast(sum(fixalloc)*100 as varchar(10))
#                                   from dbo.unplatformalloc where partPlanId = ?
#                                   group by fundcode, recordtype, InvCode, SourceCode""", self.Plan_text.get())
#                    sql_data = cursor.fetchall()
#
#                    j = 4
#                    for row in sql_data:
#                        i = 0
#                        for data in row:
#                            Label(frm1, text=data, anchor=W, font=("Times New Roman", 12)).grid(row=i + 1,
#                                                                                                column=j + 1,
#                                                                                                sticky="w")
#                            i += 1
#                        j += 1
#
#            except pyodbc.Error as Err3:
#                tkinter.messagebox.showinfo("Error", "Something Went wrong!!Please check connection or Input Details "
#                                                     "and Try again")
#                print(Err3)

    def get_client_details(self):
        if self.Plan_text.get() == '':
            tkinter.messagebox.showinfo("Error", "Please Enter Plan No(T0..)")
        else:
            try:
                column_names = ["Product", "Application Id", "Part Plan Id", "Gender", "First Name", "Name", "Surname",
                                "Country", "DOB", "Nationality", "Contact No", "Email", "Plan Issue Date", "Advisor",
                                "Start Date", "End Date", "Plan Term", "Plan Current Status"]

                cursor.execute("""Select planno, applicant, planid, gender, firstname, first1, surname, res_country, 
                               Case When Cast(dob as Varchar(10)) = '1900-01-01' Then 'N/A'
                               Else Cast(dob as Varchar(10)) End ,
                               nationality, cellno, emailaddr, issuedate, agent, 
                               Case When Cast(startdate as Varchar(10)) = '1900-01-01' Then 'N/A'
                               Else Cast(startdate as Varchar(10)) End , 
                               Case When Cast(enddate as Varchar(10)) = '1900-01-01' Then 'N/A'
                               Else Cast(enddate as Varchar(10)) End , term, 
                               Case 
                               When Cast(status as varchar(50)) = '0' Then 'Active'
                               When Cast(status as varchar(50)) = '3' Then 'Active but Ineligible'
                               When Cast(status as varchar(50)) = '7' Then 'Suspended Contributions'
                               When Cast(status as varchar(50)) = '30' Then 'Terminated'
                               When Cast(status as varchar(50)) = '31' Then 'Terminated'
                               When Cast(status as varchar(50)) = '32' Then 'Terminated'
                               Else Cast(status as varchar(50))
                               End
                               from dbo.unplatform where planid = ?""", self.Plan_text.get())
                sql_data = cursor.fetchone()
#                print(sql_data)
                if sql_data is None:
                    tkinter.messagebox.showinfo("Error", "Please check Plan No and try again")
                else:
                    root2 = Toplevel(self.master)
                    root2.geometry('1150x350+0+0')
                    root2.title('Client Details')
                    frm1 = Frame(root2, width=800, height=150, relief=RAISED, border=5)
                    frm1.grid(row=0)
#                    frm2 = Frame(root2, width=600, height=150, relief=RAISED, border=5)
#                    frm2.grid(row=0, column=7)
#                    frm3 = Frame(root2, width=400, height=150, relief=RAISED, border=5)
#                    frm3.grid(row=0, column=10)
                    Label(frm1, text="Basic Details", font=("Times New Roman", 18), fg='gray40', anchor=W).grid(row=0,
                                                                                                             column=0,
                                                                                                             sticky="w")
                    i = j = c = 0
                    for row in column_names:
                        if i <= 10:
                            Label(frm1, text=row, anchor=NW, font=("Times New Roman", 12), fg='royal blue').grid(
                                                                                                           row=i + 1,
                                                                                                           column=c,
                                                                                                           sticky="w")
                            i += 1
                            if i == 11:
                                c += 2
                        else:
                            Label(frm1, text=row, anchor=NW, font=("Times New Roman", 12), fg='royal blue').grid(
                                                                                                           row=j + 1,
                                                                                                           column=c,
                                                                                                           sticky="w")
                            j += 1
                            if j == 11:
                                i = j = 0
                                c += 2

                    i = j = 0
                    c = 1
                    for row in sql_data:
                        if row == '':
                            row = 'N/A'
                        if i <= 10:
                            Label(frm1, text=row, anchor=W, font=("Times New Roman", 12)).grid(row=i + 1, column=c,
                                                                                               sticky="w")
                            i += 1
                            if i == 11:
                                c += 2
                        else:
                            Label(frm1, text=row, anchor=W, font=("Times New Roman", 12)).grid(row=j + 1, column=c,
                                                                                               sticky="w")
                            j += 1
                            if j == 11:
                                i = j = 0
                                c += 2
                    Label(frm1, text="Contribution Details", anchor=W,
                          font=("Times New Roman", 18), fg='gray40').grid(row=0, column=7, sticky="w")
                    column_names = ["Contribution Currency", "Contribution Amount", "Employer Cont Amount", "Pay Freq",
                                    "Payment Method", "DDI Bank Name", "DDI Acc No", "DDI Reject Date"]
                    cursor.execute("""Select contribcurr, contribamt, employercontrib, 
                                   Case When Cast(payfreq as Varchar(20)) = 1 Then 'Weekly'
                                   When Cast(payfreq as Varchar(20)) = 2 Then 'Bi-Weekly'
                                   When Cast(payfreq as Varchar(20)) = 3 Then 'Semi-Monthly'
                                   When Cast(payfreq as Varchar(20)) = 4 Then 'Monthly'
                                   When Cast(payfreq as Varchar(20)) = 5 Then 'Quarterly'
                                   When Cast(payfreq as Varchar(20)) = 6 Then 'Half-Yearly'
                                   When Cast(payfreq as Varchar(20)) = 7 Then 'Annual'
                                   Else 'Unknown'
                                   End, 
                                   Case When Cast(paymethod as Varchar(20)) = 1 Then 'Bank Transfer'
                                   When Cast(paymethod as Varchar(20)) = 2 Then 'Cash'
                                   When Cast(paymethod as Varchar(20)) = 3 Then 'Cheque'
                                   When Cast(paymethod as Varchar(20)) = 4 Then 'Credit Card'
                                   When Cast(paymethod as Varchar(20)) = 5 Then 'DDI'
                                   When Cast(paymethod as Varchar(20)) = 6 Then 'Electronic Transfer'
                                   When Cast(paymethod as Varchar(20)) = 7 Then 'Standing Order'
                                   Else 'Other'
                                   End
                                   , ddibankname,
                                   ddiaccno, Case When Cast(ddirejectdate as Varchar(10)) = '1900-01-01' Then 'N/A' 
                                   Else Cast(ddirejectdate as Varchar(10)) End
                                   from dbo.unplatform where planid = ?""", self.Plan_text.get())
                    sql_data = cursor.fetchone()
#                    print(sql_data)
                    i = 0
                    for row in column_names:
                        Label(frm1, text=row, anchor=W, font=("Times New Roman", 12), fg='royal blue').grid(row=i + 1,
                                                                                                            column=7,
                                                                                                            sticky="w")
                        i += 1
                    i = 0
                    for row in sql_data:
                        if row == '':
                            row = 'N/A'
                        Label(frm1, text=row, anchor=W, font=("Times New Roman", 12)).grid(row=i + 1, column=8,
                                                                                               sticky="w")
                        i += 1

                    Label(frm1, text="Balance Details", anchor=W,
                          font=("Times New Roman", 18), fg='gray40').grid(row=0, column=10, sticky="w")
                    column_names = ["Num Of Contributions", "Last Contribution Date", "Contribution Source",
                                    "Total Amount Contributed", "Current Balance"]

                    i = 0
                    for row in column_names:
                        Label(frm1, text=row, anchor=W, font=("Times New Roman", 12), fg='royal blue').grid(row=i + 1,
                                                                                                            column=10,
                                                                                                            sticky="w")
                        i += 1
#                    Label(frm1, text="Last Contribution Date", anchor=W, font=("Times New Roman", 12)).grid(row=1,
#                                                                                                            column=8,
#                                                                                                            sticky="w")
#                    Label(frm1, text="Contribution Source", anchor=W, font=("Times New Roman", 12)).grid(row=1,
#                                                                                                         column=9,
#                                                                                                         sticky="w")
                    cursor.execute("""Select count(*), max([TRADEDATE-BR008]), [FUNDSRC-BR099], sum([AMOUNT-BR110])
                                               from ODS.OMNI_DC_TXN_HIST where PARTPLANID = ? AND
                                               [TRANCODE-BR101] = '114' AND [ACTIVITY-BR102] = 1 AND
                                               [STDUSAGECOD_3-BR303] = ''
                                               group by [FUNDSRC-BR099]""", self.Plan_text.get())
                    sql_data = cursor.fetchall()
                    #                    print(sql_data)
#                    i = 1
#                    j = 0
                    j = 10
                    for row in sql_data:
                        i = 0
                        for data in row:
                            Label(frm1, text=data, anchor=W, font=("Times New Roman", 12)).grid(row=i+1,
                                                                                                column=j+1,
                                                                                                sticky="w")
                            i += 1
                        j += 1

            except pyodbc.Error as Err2:
                tkinter.messagebox.showinfo("Error", "Something Went wrong!!Please check connection or Input Details "
                                                     "and Try again")
                print(Err2)


root = Tk()
Omni_gui = Window(root)

root.mainloop()
