from os import environ, makedirs, listdir, rmdir
from os.path import exists, abspath
from subprocess import Popen

from collections import OrderedDict
from decimal import Decimal, getcontext, ROUND_HALF_EVEN
from datetime import datetime
from textwrap import fill, wrap

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import *

import pypandoc
from mailmerge import MailMerge
from openpyxl import load_workbook

from matplotlib import use, rcParams
use("Agg")
from matplotlib import pyplot

CHARTS_DIR = "Charts"
BILLS_DIR = "Bills"

environ.setdefault("PYPANDOC_PANDOC", "C:/Program Files (x86)/Pandoc")
getcontext().rounding = ROUND_HALF_EVEN  # so it's not locale-specific
# Matplotlib settings
rcParams.update({'xtick.labelsize': 'small',
                 'figure.dpi': 150.0,
                 'figure.figsize': [8/3, 2.0]})
# pyplot.autoscale(False)


class Tenant:
    def __init__(self, unit_no, document):
        self.unit_no = unit_no
        self.document = document
        self.chart_path = "{}/unit_{}_history.png".format(CHARTS_DIR, self.unit_no)
        self.bill_path = "{}/unit_{}_{}_bill.docx".format(
            BILLS_DIR, self.unit_no, BILLS_DIR[6:])  # formatted period-name
        self.four_recent_bills = []  # type: [Decimal]
        self.current_reading = self.prev_reading = Decimal('0.00')
        self.name = self.account_no = self.meter_no = self.service_addr = None
        self.billing_addr = self.billing_city = None
        self.billing_prov = self.billing_postal = None
        self.prev_balance = self.consumption_total = self.amount_due = Decimal('0.00')
        # prev_balance left at 0 for now, can be used later

    def get_addr_info(self, row):
        """
        Fills in this tenant's name and address information from
        the given row from the TenantInfo sheet of the Excel document.

        :param List[Cell] row: this tenant's row in the TenantInfo sheet
        :return: None
        """
        self.meter_no = str(row[1].value)
        self.name = row[2].value
        self.account_no = '-'  # not used right now, but needed for template
        self.service_addr = row[3].value
        self.billing_addr = row[7].value
        self.billing_city = row[8].value
        self.billing_prov = row[9].value
        self.billing_postal = row[10].value

    def calculate_bills(self, row, index, rate):
        """
        Calculates the current bill and up to 3 previous bills
        for this tenant.

        :param List[Cell] row: the Excel row for this tenant
        :param int index: the index to the current period to calculate
        :param str rate: the current month's billing rate (as a string)
        :return: None
        """
        print([c.value for c in row])
        self.current_reading = str(row[index].value)
        self.prev_reading = str(row[index - 1].value)
        self.four_recent_bills.append(Decimal(self.current_reading) -
                                      Decimal(self.prev_reading))
        print(rate, type(rate))
        self.consumption_total = Decimal(self.four_recent_bills[0] *
                                         Decimal(rate))
        self.amount_due = str(round(self.prev_balance + self.consumption_total))
        self.consumption_total = str(round(self.consumption_total))

        count = 0
        # print("calculating bills:")
        while count < 3 and index > 2:
            index -= 1
            count += 1
            # print(index, count)
            self.four_recent_bills.append(
                Decimal(row[index].value) - Decimal(row[index - 1].value))
        while count < 3:
            count += 1
            # print(count)
            self.four_recent_bills.append(0)

    def generate_chart(self, index, periods):
        """

        :type index: int
        :type periods: OrderedDict
        :return: None
        """
        def autolabel(rects, axes):
            """
            Attach a text label above each bar displaying its height
            Thanks to the API and Lindsey Kuper at composition.al.
            """
            (y_bottom, y_top) = axes.get_ylim()
            y_height = y_top - y_bottom
            for rect in rects:
                height = rect.get_height()
                if height > 0:
                    label_position = height + (y_height * 0.05)
                    axes.text(rect.get_x() + rect.get_width() / 2.,
                              label_position, str(round(height, 2)),
                              ha='center', va='bottom')

        bills = [round(b) for b in self.four_recent_bills[::-1]]  # reverse
        print(bills)
        rng = range(4)
        fig, ax = pyplot.subplots()  # type: Figure, Axes
        # ax.set_autoscaley_on(False)

        ax.set_xticks(rng)
        labels = []
        i = -2
        periods = list(periods.items())
        while i >= -5:
            text = fill(periods[index + i][0], 10) if index + i >= 0 else ''
            labels.insert(0, text)
            i -= 1
        ax.set_xticklabels(labels)
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=90)

        ax.axes.get_yaxis().set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['right'].set_visible(False)

        chart = ax.bar(rng, bills, 1/1.5, color="blue")
        chart[3].set_color("cyan")
        autolabel(chart, ax)
        # fig.tight_layout()
        fig.subplots_adjust(bottom=0.4)
        # print(rcParams['figure.subplot.bottom'])
        # print("ylim:", pyplot.ylim())
        # print("fig extent:", fig.get_window_extent())
        # print("ax extent:", ax.get_window_extent())
        fig.savefig(self.chart_path)
        # fig.show()
        pyplot.close(fig)

    def generate_bill(self, start, end, days, rate, due_date):
        # d = dict(Name=self.name,
        #          AccountNo=self.account_no,
        #          MeterNo=self.meter_no,
        #          ServiceAddr=self.service_addr,
        #          BillingAddr=self.billing_addr,
        #          BillingCity=self.billing_city,
        #          BillingProv=self.billing_prov,
        #          BillingPostal=self.billing_postal,
        #          StartDate=start,
        #          EndDate=end,
        #          NumberDays=days,
        #          PrevReading=self.prev_reading,
        #          CurrentReading=self.current_reading,
        #          AmountConsumed=str(self.four_recent_bills[0]),  # in m^3
        #          DueDate=due_date,
        #          ServiceRate=rate,
        #          PrevBalance=str(self.prev_balance),
        #          TotalConsumption=self.consumption_total,  # in $
        #          TotalDue=self.amount_due)
        # print("Non-strings:")
        # for x in d:
        #     if not isinstance(d[x], str):
        #         print(x)
        self.document.merge(Name=self.name,
                            AccountNo=self.account_no,
                            MeterNo=self.meter_no,
                            ServiceAddr=self.service_addr,
                            BillingAddr=self.billing_addr,
                            BillingCity=self.billing_city,
                            BillingProv=self.billing_prov,
                            BillingPostal=self.billing_postal,
                            StartDate=start,
                            EndDate=end,
                            NumberDays=days,
                            PrevReading=self.prev_reading,
                            CurrentReading=self.current_reading,
                            AmountConsumed=str(self.four_recent_bills[0]),  # in m^3
                            DueDate=due_date,
                            ServiceRate=rate,
                            PrevBalance=str(self.prev_balance),
                            TotalConsumption=self.consumption_total,  # in $
                            TotalDue=self.amount_due,
                            ChartName=abspath(self.chart_path))
        self.document.write(self.bill_path)


class BillGenerator(Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Submeter Bill Generator")

        self.template = StringVar(self)
        self.periods = self.period_index = self.data = self.contacts = \
            self.start_date = self.end_date = self.no_days = \
            self.rate = self.due_date = None
        self.tenants = []

        self.run()

    def run(self):
        period_menu = None
        option = StringVar(self)
        datafile = StringVar(self)

        def get_data_file(*args):
            fn = filedialog.askopenfilename(
                filetypes=[("Excel files", ".xlsx")], parent=self)
            if not fn:
                messagebox.showerror("Opening File",
                                     "Cannot open file. Please try again.",
                                     parent=self)
            else:
                datafile.set(fn)

        def read_data_file(*args):
            ok_button.pack_forget()
            message.configure(text="Data file open and reading. "
                                   "Please wait.")
            wb = load_workbook(filename=datafile.get(), data_only=True)
            self.data = wb["DataEntry"]
            self.contacts = wb["TenantInfo"]
            self.periods = OrderedDict((c.value, i+2) for (i, c)
                                       in enumerate(self.data[1][2:])
                                       if c.value is not None)

            # Choose from available billing periods
            nonlocal period_menu
            last = self.periods.popitem()
            self.periods[last[0]] = last[1]
            last = last[0]
            period_menu = OptionMenu(mainframe, option, last,
                                     *self.periods.keys())
            message.configure(text="Please select a billing period_index\n"
                                   "to produce bills for:")
            ok_button.configure(text="OK", command=select_period)
            period_menu.grid(row=2, column=1)
            ok_button.pack(pady=10, padx=10)

        def select_period(*args):
            print("start of select_period")
            self.period_index = self.periods[option.get()]
            if any(len(x) > 10 for x in wrap(option.get(), 10)):
                messagebox.showwarning("Format Warning",
                                       "Warning: This period name '%s' "
                                       "contains word(s) \nlonger than 10 "
                                       "characters. This may result in a "
                                       "formatting error.\nChange the period "
                                       "name in the Excel file in accordance "
                                       "with \nthe naming standard, or proceed "
                                       "with caution." % option.get(),
                                       parent=self)
            print(self.period_index)

            # Add billing period to Charts and Bills folder names
            global CHARTS_DIR, BILLS_DIR
            CHARTS_DIR += '/' + option.get()
            CHARTS_DIR = CHARTS_DIR.replace('- ', '-').replace(' -', '-')
            CHARTS_DIR = CHARTS_DIR.replace(' ', '_')
            BILLS_DIR += '/' + option.get()
            BILLS_DIR = BILLS_DIR.replace('- ', '-').replace(' -', '-')
            BILLS_DIR = BILLS_DIR.replace(' ', '_')
            print(CHARTS_DIR, BILLS_DIR)

            period_menu.grid_forget()
            message.configure(text="Please select a Word file to use\n"
                                   "as a template for the bills:")
            self.template.trace('w', use_template_file)
            ok_button.configure(text="Choose File",
                                     command=get_template_file)
            print("end of select_period")

        def get_template_file(*args):
            fn = filedialog.askopenfilename(
                filetypes=[("Word files", ".docx")], parent=self)
            if not fn:
                messagebox.showerror("Opening File",
                                     "Cannot open file. Please try again.",
                                     parent=self)
            else:
                self.template.set(fn)

        def use_template_file(*args):
            ok_button.pack_forget()
            message.configure(text="Template file opened. Generating bills.\n"
                                   "Please wait, as this may take some time.")
            self.template = self.template.get()
            get_period_info()
            initialise_tenants()
            print(CHARTS_DIR, BILLS_DIR)
            print("calling generate_charts")
            generate_charts()
            print(self.tenants[0].document.get_merge_fields())
            print("calling generate_bills")
            generate_bills()

        def get_period_info():
            period = self.data[chr(ord('A') + self.period_index)]
            # Already datetime objects (thanks, openpyxl)
            # self.start_date = datetime.strptime(period[1].value, "%Y-%m-%d")
            # self.end_date = datetime.strptime(period[2].value, "%Y-%m-%d")
            self.start_date = period[1].value  # type: datetime
            self.start_date = self.start_date.strftime("%b %d/%y")  # Jan 31/17
            self.end_date = period[2].value.strftime("%b %d/%y")
            self.no_days = str(period[3].value)
            self.rate = str(period[4].value)
            self.due_date = period[5].value.strftime("%b %d/%y")
            print(self.start_date, self.end_date, self.no_days, self.rate, self.due_date)

        def initialise_tenants():
            for i, unit_row in enumerate(self.data.iter_rows(min_row=7)):
                unit = unit_row[0].value
                if unit is None:
                    break
                t = Tenant(unit_no=str(unit), document=MailMerge(self.template))
                contact_row = self.contacts[i+2]
                assert contact_row[0].value == unit, \
                    "Mismatch in the tenant order of DataEntry and " \
                    "TenantInfo sheets ({} != {})".format(
                        contact_row[0].value, unit)
                t.get_addr_info(contact_row)
                t.calculate_bills(unit_row, self.period_index, self.rate)
                self.tenants.append(t)

        def generate_charts():
            assert not exists(CHARTS_DIR), \
                "'{}' folder already exists.".format(CHARTS_DIR)
            makedirs(CHARTS_DIR)
            for tenant in self.tenants:
                tenant.generate_chart(self.period_index, self.periods)

        def generate_bills():
            assert not exists(BILLS_DIR), \
                "'{}' folder already exists.".format(BILLS_DIR)
            makedirs(BILLS_DIR)
            for tenant in self.tenants:
                tenant.generate_bill(self.start_date, self.end_date,
                                     self.no_days, self.rate, self.due_date)
            print("done!")
            close()

        def close():
            path = abspath(BILLS_DIR)
            message.configure(text="Bill generation complete!\n"
                                   "Bills have been stored in %s.\n"
                                   "Press Open Folder to open this folder \n"
                                   "in Explorer, or press Exit to close.")
            ok_button.configure(text="Open Folder",
                                command=lambda: Popen("explorer " + path))
            exit_button.configure(command=lambda: self.quit())
            exit_button.pack(padx=10, pady=10)
            self.quit()

        # Add a grid
        mainframe = Frame(self)
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)
        mainframe.pack(padx=100, pady=50)

        # Select data file
        message = Label(mainframe, text="Select a data file:", justify=CENTER)
        message.grid(row=1, column=1)
        exit_button = Button(self, text="Exit")
        ok_button = Button(self, text="Choose File", command=get_data_file)
        datafile.trace('w', read_data_file)
        ok_button.pack(padx=10, pady=10)


if __name__ == "__main__":
    try:
        root = BillGenerator()
        root.mainloop()
    except Exception as e:
        if exists(CHARTS_DIR) and not listdir(CHARTS_DIR):
            rmdir(CHARTS_DIR)
        if exists(BILLS_DIR) and not listdir(BILLS_DIR):
            rmdir(BILLS_DIR)
        raise e
