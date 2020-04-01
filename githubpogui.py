
from tkinter import *
import tkinter as tk
from tkinter import filedialog

def newcompany(): #for future companies input
	pass

def jojorun(): #company A when run
	import re
	import pdfplumber
	import pandas as pd
	from collections import namedtuple
	file = filedialog.askopenfilename(filetypes = ((("PDF", "*.pdf"), )))

	def loadtemplate():
		dirname = filedialog.askdirectory(parent=window, initialdir="/", title = 'Please select a destination')

		Line = namedtuple('Line', 'Company PO_Number IDx Item_No Variant_Code Description Age In_Years_Or_Months Qty Unit_Cost Disc Amount')
		#Finding the line by using Item No. formate (using regex)
		line_re = re.compile(r'\b[A-Z]{2}\d{3}\b') 
		company_re = re.compile(r'(Company A)')
		purchase_re = re.compile(r'(\b[A-Z]{4}\d{8}\b)')
		#address_re is not being used right now
		address_re = re.compile(r'(DELIVERY ADDRESS)')

		lines = []
		total_check = 0

		with pdfplumber.open(file) as pdf:
			pages = pdf.pages
			for page in pdf.pages:
				text=page.extract_text()
				for line in text.split('\n'):
					purchase = purchase_re.search(line)
					comp = company_re.search(line)
					if comp:
						company_name = comp.group(1)
					elif purchase:
						po_no = purchase.group(1)
					elif line_re.search(line):
						parts = line.split()
						items = parts[:3] + [" ".join(parts[3:-6])] + parts[-6:]
						lines.append(Line(company_name, po_no, *items))


		destinationkey = dirname + '/' + po_no + '.xlsx'			
		df = pd.DataFrame(lines)
		destinationkey = destinationkey.replace("/", "\\")
		df.to_excel(destinationkey, sheet_name='PO', index = False)
		print(destinationkey)

	output_button = tk.Button(window, text="Select Destination", command = loadtemplate, width=10)
	output_button.grid(row=0, column=4)
	
window=Tk()
window.title("PO Reader Alpha")#Title

entry_company = Entry(window, text="Enter Company Name")
#user can input either "company a" or just "a"
entry_company.grid(row=0, column=0)

def quit():
	window.quit() #for closing programme function
        
def company_input():
	input_company = entry_company.get()
	#makes it so that the input is not case sensitive
	x = input_company.lower()
	if x == "company a" or x== "a":
		import_button = tk.Button(window, text="Import PDF & Convert", command=jojorun)
		import_button.grid(row=0, column = 2)
		newlabel = tk.Label(window, text="Company A is selected")
		newlabel.grid(row=1, column=0)
	
enter_button = tk.Button(window, text="Enter", command=company_input)
enter_button.grid(row=0, column=1)
info = tk.Label(window, text="PO Reader v.Alpha")
info.grid(row=4, column=1)
close_button = tk.Button(window, text="Close", command=quit)
close_button.grid(row=4, column=0)

#currently servers no function but is a placeholder for a future drag and drop mechanic
list1 = Listbox(window, height=6, width=35)
list1.grid(row=3, column=0)
window.mainloop()