import tkinter as tk
from tkinter import ttk


class InputWindow(tk.Toplevel):

    def __init__(self, *args, callback=None, **kwargs):
        super().__init__(*args, **kwargs)
        # callback is a function that this window will call
        # with the entered name as an argument once the button
        # has been pressed.
        self.callback = callback
        self.config(width=300, height=90)
        # Disable the button for resizing the window.
        self.resizable(0, 0)
        self.title("Enter Your Name")
        self.entry_name = ttk.Entry(self)
        self.entry_name.place(x=20, y=20, width=260)
        self.button_done = ttk.Button(
            self,
            text="Done!",
            command=self.button_done_pressed
        )
        self.button_done.place(x=20, y=50, width=260)
        self.focus()
        self.grab_set()

    def button_done_pressed(self):
        # Get the entered name and invoke the callback function
        # passed when creating this window.
        self.callback(self.entry_name.get())
        # Close the window.
        self.destroy()


class MainWindow(tk.Tk):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.config(width=400, height=300)
        self.title("Main Window")
        self.button_request_name = ttk.Button(
            self,
            text="Request name",
            command=self.request_name
        )
        self.button_request_name.place(x=50, y=50)
        self.label_name = ttk.Label(
            self,
            text="You have not entered your name yet."
        )
        self.label_name.place(x=50, y=150)

    def request_name(self):
        # Create the child window and pass the callback
        # function by which we want to receive the entered
        # name.
        self.ventana_nombre = InputWindow(
            callback=self.name_entered
        )

    def name_entered(self, name):
        # This function is invoked once the user presses the
        # "Done!" button within the secondary window. The entered
        # name will be in the "name" argument.
        self.label_name.config(
            text="Your name is: " + name
        )


main_window = MainWindow()
main_window.mainloop()