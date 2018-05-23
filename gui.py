from tkinter import Tk, Label, Button

class kayaGUI:
    def __init__(self, master):
        self.master = master
        master.title("Manejo de inventarios Kaya Unite")

        self.label = Label(master, text="This is our first GUI!")
        self.label.pack()

        self.greet_button = Button(master, text="Greet", command=self.greet)
        self.greet_button.pack()

        self.close_button = Button(master, text="Close", command=master.quit)
        self.close_button.pack()

    def greet(self):
        print("Greetings!")

root = Tk()
my_gui = kayaGUI(root)
root.mainloop()
