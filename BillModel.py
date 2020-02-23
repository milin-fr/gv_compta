class BillModel:

    row_placement = ""
    work_type = ""  #1
    company_name = ""  #2
    comment = ""  #3
    start_date = ""  #4
    end_date = ""  #5
    price = ""  #6
    payment_status = ""  #7
    work_type_budget = ""
    work_type_budget_leftover = ""
    total_budget = ""
    total_budget_leftover = ""
    excel_file_name = ""

    def get_row_placement(self):
        return self.row_placement

    def set_row_placement(self, row_placement):
        self.row_placement = row_placement


    def get_work_type(self):
        return self.work_type

    def set_work_type(self, work_type):
        self.work_type = work_type


    def get_company_name(self):
        return self.company_name

    def set_company_name(self, company_name):
        self.company_name = company_name


    def get_comment(self):
        return self.comment

    def set_comment(self, comment):
        self.comment = comment


    def get_start_date(self):
        return self.start_date

    def set_start_date(self, start_date):
        self.start_date = start_date


    def get_end_date(self):
        return self.end_date

    def set_end_date(self, end_date):
        self.end_date = end_date


    def get_price(self):
        return self.price

    def set_price(self, price):
        self.price = price


    def get_payment_status(self):
        return self.payment_status

    def set_payment_status(self, payment_status):
        self.payment_status = payment_status


    def get_work_type_budget(self):
        return self.work_type_budget

    def set_work_type_budget(self, work_type_budget):
        self.work_type_budget = work_type_budget


    def get_work_type_budget_leftover(self):
        return self.work_type_budget_leftover

    def set_work_type_budget_leftover(self, work_type_budget_leftover):
        self.work_type_budget_leftover = work_type_budget_leftover


    def get_total_budget(self):
        return self.total_budget

    def set_total_budget(self, total_budget):
        self.total_budget = total_budget

    def get_total_budget_leftover(self):
        return self.total_budget_leftover

    def set_total_budget_leftover(self, total_budget_leftover):
        self.total_budget_leftover = total_budget_leftover

    def get_excel_file_name(self):
        return self.excel_file_name

    def set_excel_file_name(self, excel_file_name):
        self.excel_file_name = excel_file_name

"""
    def set_excel_name(self):
        self.excel_file_name = "GV compta " + self.work_type + ".xlsx"
"""