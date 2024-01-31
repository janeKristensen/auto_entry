
from tkinter import Tk, ttk, filedialog as td

isCancel = False


def browse_file():
    global chosen_file 

    chosen_file = td.askopenfilename(
        multiple=True,
        filetypes=[("Excel files", ".xls .xlsx")],
        )               
    if chosen_file:
        root.destroy()

def browse_cancel():
    global isCancel

    isCancel = True
    root.destroy()


def get_filename():

    try:
        global root
        root = Tk()
        root.title("Load File")

        # Set dimensions of the screen
        width = 300
        height = 150
    
        # Get screen width and height
        width_screen = root.winfo_screenwidth() 
        height_screen = root.winfo_screenheight() 

        # Calculate x and y coordinates for the Tk root window
        x = (width_screen/2) - (width/2)
        y = (height_screen/2) - (height/2)

        # Set the dimensions of the screen 
        # and where it is placed
        root.geometry('%dx%d+%d+%d' % (width, height, x, y))

        # Create a frame
        frm = ttk.Frame(root, padding=30)
        frm.grid()

        # Create labels
        lbl = ttk.Label(frm, text='Select a reference request form to open.')
        lbl2 = ttk.Label(frm, text='File type must be .xls or .xlsx')

        # Create buttons
        btn_browse = ttk.Button(frm, text='Browse files',command=browse_file)   
        btn_cancel = ttk.Button(frm, text='Cancel', command=browse_cancel)

        # Layout all the widgets
        lbl.grid(row=0, column=0, padx=5, pady=5, columnspan=2)  
        lbl2.grid(row=1, column=0, padx=5, pady=5, columnspan=2)   
        btn_browse.grid(row=2, column=0, padx=5, pady=5)
        btn_cancel.grid(row=2, column=1, padx=5, pady=5)

        # Display the frame
        root.attributes('-topmost', True)
        root.focus_force()
        root.update()
        frm.pack()
        root.mainloop()

        if isCancel:
            return None
        else:
            return chosen_file
        
    except Exception as e:
        print({str(e)})
        pass
    
 







