from tkinter import *
from tkinter import ttk
import re
from settings import PROGRAM_NAME,MAX_ORDER_ROWS, EXCEL_FILE, FILE_LOCATION
from excel_functions import *
from helper_func import check_if_already_exists

wb = load_workbook(filename = f"{FILE_LOCATION}/{EXCEL_FILE}")
database = "Test Database"
history = "Test historikk"
#https://stackoverflow.com/questions/66412903/remove-the-default-0-in-in-integer-entry-in-tkinter
#how to do the entry as a stringvar and output. aka just cast it to int. that will do it.
def find_descriptions(*args):
    order_numbers = get_all_saved_unique_ordernumbers(wb,history,"g")
    if check_if_already_exists(order_numbers,order_number.get()) == False:
        #TODO if invalid, make the submit button's state unusable?
        valid_invalid_text.set("invalid")
        return
    else:
        valid_invalid_text.set("Valid")
    for i in range(1,MAX_ORDER_ROWS+1):
        value = product_numbers[f"product_number_{i}"].get()

        product_description = find_product_description(wb, database, value)

        if product_description == None:
            descriptions[f"item_{i}"].set("ERROR, COULD NOT FIND PRODUCT")
        else:
            descriptions[f"item_{i}"].set(product_description)

def reset_form(*args):
    #i have to reset the form somehow. do i put the creation of the form into a function, call that function as normal,
    #then if i need to reset the form i call it in the reset form button?
    # button is pressed->order number entry+ label is reset-> product_numbers, descriptions, product_amounts is reset to empty dicts
    # then entries and labels for all products is reset.
    #no, dont nessesarily need to empty dictionary then re-create the whole fields. i just need to set the fields to 0(for intvar)
    #  or ""(for stringvar). what is saved in the dictionary is updated on the go as the change is happening
    #Should i make the creation of the ordernum + products fields a separate function?
    #when i click submit the sameThing should happen


    # orders packed today is NOT reset. need a function to check todays date, and orders packed today?
    pass
def check_num(newval):
    return re.match('^[0-9]*$', newval) is not None and len(newval) <= 10


root = Tk()
root.title(f"{PROGRAM_NAME}")

check_num_wrapper = (root.register(check_num), '%P')

mainframe = ttk.Frame(root,padding=(400,3,3,12))
mainframe.grid(column=0,row=0, sticky=(N,S,E,W))

# test=ttk.Label(mainframe,text="test").grid(column=0,row=0)
#Title
ttk.Label(mainframe, text="pack orders:",padding=(5,5,5,5)).grid(column=1,row=1, sticky=(N,S,E,W))

#order number entry and text
ttk.Label(mainframe, text="order number:").grid(column=1, row=2,sticky=E)
order_number = IntVar()
valid_invalid_text = StringVar()
order_number_entry = ttk.Entry(mainframe,textvariable=order_number,validate="key",validatecommand=check_num_wrapper)
order_number_entry.grid(column=2, row=2,sticky=(W,E))
ttk.Label(mainframe,textvariable=valid_invalid_text).grid(column=3, row=2,sticky=W)


#buttons, added to a different frame.
buttonframe = ttk.Frame(root,padding=(1,71,300,30))
buttonframe.grid(column=1,row=0,sticky=(N))
ttk.Button(buttonframe,text="find descriptions",command=find_descriptions,width=20).grid(column=0, row=0,rowspan=2,ipady=11)
ttk.Button(buttonframe,text="reset form",command=None,width=20).grid(column=0,row=2,rowspan=2,ipady=11)
ttk.Button(buttonframe,text="submit", command=None,width=20).grid(column=0, row=100,rowspan=2,ipady=11,pady=50)


#buttons todo: buttons makes the padding around the entry wonkey, fix by making a 2nd frame? fixed above
# ttk.Button(mainframe,text="find descriptions",command=find_descriptions,width=20).grid(column=5, row=4,rowspan=2,ipady=11)
# ttk.Button(mainframe,text="reset form",command=None,width=20).grid(column=5,row=6,rowspan=2,ipady=11)
# ttk.Button(mainframe,text="submit", command=None,width=20).grid(column=5, row=100,rowspan=2,ipady=11)


orders_sent_today_count = 0
ttk.Label(mainframe, text="orders sent today:").grid(column=2, row=100, sticky=(N,E),pady=3)
orders_sent_today = ttk.Label(mainframe, text=orders_sent_today_count)
orders_sent_today.grid(column=3, row=100,sticky=W,pady=5)

#product number and amount grid
ttk.Label(mainframe,text="product number:").grid(column=1, row=3,sticky=E)
ttk.Label(mainframe, text="description").grid(column=2, row=3)
ttk.Label(mainframe,text="amount").grid(column=3,row=3,sticky=W)

product_numbers= {

}
descriptions = {

}
product_amounts = {

}
for i in range(1,MAX_ORDER_ROWS+1):
    product_numbers[f"product_number_{i}"] = IntVar()
    descriptions[f"item_{i}"] = StringVar()
    product_amounts[f"amount_{i}"] = IntVar()
    ttk.Entry(mainframe,textvariable=product_numbers[f"product_number_{i}"],validate="key", validatecommand=check_num_wrapper).grid(column=1,row=3+i)
    ttk.Label(mainframe,textvariable=descriptions[f"item_{i}"]).grid(column=2, row=3+i,sticky=W)
    ttk.Entry(mainframe,textvariable=product_amounts[f"amount_{i}"],validate="key", validatecommand=check_num_wrapper).grid(column=3, row=3+i)


#extra config

root.columnconfigure(0,weight=1)
root.rowconfigure(0,weight=1)
mainframe.columnconfigure(2, weight=1)

# for child in mainframe.winfo_children():
#     child.grid_configure(padx=2, pady=2)
root.mainloop()

# product_number_1 = IntVar()
# description_1 = StringVar()
# amount_1 = IntVar()
# product_number_1_entry = ttk.Entry(mainframe,textvariable=product_number_1).grid(column=1,row=4)
# ttk.Label(mainframe,textvariable=descriptions["item_1"]).grid(column=2, row=4,sticky=W)
# product_number_1_amount = ttk.Entry(mainframe,textvariable=amount_1).grid(column=3, row=4)

# product_number_2 = IntVar()
# description_2 = StringVar()
# amount_2 = IntVar()
# product_number_2_entry = ttk.Entry(mainframe,textvariable=product_number_2).grid(column=1,row=5)
# ttk.Label(mainframe,textvariable=description_2).grid(column=2, row=5)
# product_number_2_amount = ttk.Entry(mainframe,textvariable=amount_2).grid(column=3, row=5)

# product_number_3 = IntVar()
# description_3 = StringVar()
# description_3.set("item3")
# amount_3 = IntVar()
# product_number_3_entry = ttk.Entry(mainframe,textvariable=product_number_3).grid(column=1,row=6)
# ttk.Label(mainframe,textvariable=description_3).grid(column=2, row=6)
# product_number_3_amount = ttk.Entry(mainframe,textvariable=amount_3).grid(column=3, row=6)

# product_number_4 = IntVar()
# description_4 = StringVar()
# amount_4 = IntVar()
# product_number_4_entry = ttk.Entry(mainframe,textvariable=product_number_4).grid(column=1,row=7)
# ttk.Label(mainframe,textvariable=description_4).grid(column=2, row=7)
# product_number_4_amount = ttk.Entry(mainframe,textvariable=amount_4).grid(column=3, row=7)

# product_number_5 = IntVar()
# description_5 = StringVar()
# amount_5 = IntVar()
# product_number_5_entry = ttk.Entry(mainframe,textvariable=product_number_5).grid(column=1,row=8)
# ttk.Label(mainframe,textvariable=description_5).grid(column=2, row=8)
# product_number_5_amount = ttk.Entry(mainframe,textvariable=amount_5).grid(column=3, row=8)

# product_number_6 = IntVar()
# description_6 = StringVar()
# amount_6 = IntVar()
# product_number_6_entry = ttk.Entry(mainframe,textvariable=product_number_6).grid(column=1,row=9)
# ttk.Label(mainframe,textvariable=description_6).grid(column=2, row=9)
# product_number_6_amount = ttk.Entry(mainframe,textvariable=amount_6).grid(column=3, row=9)

# product_number_7 = IntVar()
# description_7 = StringVar()
# amount_7 = IntVar()
# product_number_7_entry = ttk.Entry(mainframe,textvariable=product_number_7).grid(column=1,row=10)
# ttk.Label(mainframe,textvariable=description_7).grid(column=2, row=10)
# product_number_7_amount = ttk.Entry(mainframe,textvariable=amount_7).grid(column=3, row=10)

# product_number_8 = IntVar()
# description_8 = StringVar()
# amount_8 = IntVar()
# product_number_8_entry = ttk.Entry(mainframe,textvariable=product_number_8).grid(column=1,row=11)
# ttk.Label(mainframe,textvariable=description_8).grid(column=2, row=11)
# product_number_8_amount = ttk.Entry(mainframe,textvariable=amount_8).grid(column=3, row=11)

# product_number_9 = IntVar()
# description_9 = StringVar()
# amount_9 = IntVar()
# product_number_9_entry = ttk.Entry(mainframe,textvariable=product_number_9).grid(column=1,row=12)
# ttk.Label(mainframe,textvariable=description_9).grid(column=2, row=12)
# product_number_9_amount = ttk.Entry(mainframe,textvariable=amount_9).grid(column=3, row=12)

# product_number_10 = IntVar()
# description_10 = StringVar()
# amount_10 = IntVar()
# product_number_10_entry = ttk.Entry(mainframe,textvariable=product_number_10).grid(column=1,row=13)
# ttk.Label(mainframe,textvariable=description_10).grid(column=2, row=13)
# product_number_10_amount = ttk.Entry(mainframe,textvariable=amount_10).grid(column=3, row=13)

