# submeter-bill-generator
_Generate utility bills for several tenants of one building that is submetered, using a template._

This program suite is designed to speed up the process of producing beautifully-designed and accurate submeter bills in several steps:

  1. (Command line only) Downloads water usage data stored on an online database (such as at [submetersolutions.com](http://submetersolutions.com/)) as CSV files. 
  2. (Graphical interface) Calculates water usage and amount owed from the given Excel document 
     (which includes downloaded data, tenant information, and monthly rate in dollars). Currently only processes one month at a time.
  3. GUI also takes a Microsoft Word (2007 or later) template, with [jinja2-style tags](http://docxtpl.readthedocs.io/en/latest/#jinja2-like-syntax), 
     which is used to generate the final bills. (See [below](#word-document-tags).)
  4. You may want to run [merge_bills.sh](merge_bills.sh) on the resulting PDF files to merge them, 
     because there is one PDF produced for each tenant. You can use the bash script (or other PDF editor) to merge a 
     month's worth of bill PDFs into one PDF.
     
## Requirements

* Microsoft Windows (required to run VBScript)
* Python 3.x
* Python packages:
  * Tkinter (should be included in your Python installation)
  * [openpyxl](https://openpyxl.readthedocs.io/en/default/)
  * [docxtpl](http://docxtpl.readthedocs.io/en/latest/)
  * [Matplotlib](https://matplotlib.org/)
* (Optional) Bash terminal with Ghostscript installed 
  (to run merge_bills.sh, but this can also be done with Adobe Acrobat or other PDF editing software)

## Word Document Tags
These tags in the Word document should be enclosed in double brace brackets, like this: `{{ TagName }}`

* ServiceAddr (submeter address including unit number)
* Name (tenant name)
* PrevBalance (previous bill balance)
* Cons (current billing's consumption in $)
* TotalDue (current amount due as of this billing)
* StartDate (billing period start date)
* EndDate (billing period end date)
* NumDays (length of billing period)
* PRead (previous submeter reading in m^3)
* CRead (current submeter reading in m^3)
* AmCons (this billing's consumption in m^3)
* MeterNo (this tenant's submeter's ID/number)
* Rate (this month's billing rate, in $/m^3)
* BillingName (name of person/company receiving bill)
* BillingAddr, BillingCity, BillingProv, BillingPostal (address info of person/company receiving bill)
* DueDate (when this bill is due)
* AccountNo (account number)
* Chart (used for a Matplotlib image file of recent months' usage)
