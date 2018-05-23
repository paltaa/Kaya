from tkinter import *
from tkinter import filedialog
root = Tk()



def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    print(filename)

def next():
    global window
    print("Siguiente")

buttonBrowse = Button(text="Buscar Directorio", command=browse_button)
title = Label(root, text= "Manejo de inventarios KayaUnite")
config=Label(root, text="Kaya more Faya")
browseLabel= Label(root, text="Directorio de trabajo")


saleLabel=Label(root, text="Nombre archivo ventas")
saleName=Entry(root)
saleSheetName=Label(root, text= "Hoja de trabajo")
saleSheetEntry=Entry(root)


stockaLabel=Label(root, text="Nombre archivo stock inicial")
stockaName=Entry(root)
stockaSheetName=Label(root,text="Hoja de trabajo")
stockaSheetEntry=Entry(root)

stockbLabel=Label(root, text= "Nombre archivo stock anterior")
stockbName=Entry(root)
stockbSheetName=Label(root,text="Hoja de trabajo")
stockbSheetEntry=Entry(root)

buttonNext=Button(text="Siguiente", command=next)
#browseText= Label(root, text=folder_path)




#grid row column
title.grid(row=0)
config.grid(row=1)
browseLabel.grid(row=2, column=0)
#browseText.grid(row=2, column= 1)
buttonBrowse.grid(row=2, column=2)
saleLabel.grid(row=3, column=0)
saleName.grid(row=3, column=1)
saleSheetName.grid(row=3,column=2)
saleSheetEntry.grid(row=3, column=3)

stockaLabel.grid(row=4, column=0)
stockaName.grid(row=4, column=1)
stockaSheetName.grid(row=4, column=2)
stockaSheetEntry.grid(row=4, column=3)

stockbLabel.grid(row=5, column=0)
stockbName.grid(row=5, column=1)
stockbSheetName.grid(row=5, column=2)
stockbSheetEntry.grid(row=5, column=3)
root.mainloop()
