import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook
from datetime import datetime
import os

class MSnapApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("M-SNAP Form")
        self.geometry("600x650")
        self.resizable(False, False)

        self.dropdown_fields_common = {
            "Stray?": ["Yes", "No", "N/A"],
            "Older than 5 years?": ["Yes", "No", "N/A"],
            "Dog weight range?": ["Yes", "No", "N/A"],
            "Special": ["Yes", "No", "N/A"]
        }
        self.gender_field = {"Gender": ["Male", "Female", "N/A"]}

        self.init_data()
        self.build_frames()
        self.show_frame(0)

    def init_data(self):
        self.frames = []
        self.current_frame = 0
        self.add_another_pet_flags = {}

        self.data = {
            "Recipient Information": {
                "Mail to Name": tk.StringVar(),
                "Day phone(s)": tk.StringVar(),
                "Street, Apt # OR PO Box": tk.StringVar(),
                "City": tk.StringVar(),
                "Zip": tk.StringVar(),
                "Within Morg city limits?": tk.StringVar(),
                "POB only: closest town": tk.StringVar(),
                "How did you hear about M-SNAP?": tk.StringVar(),
                "Name on voucher": tk.StringVar()
            }
        }

        for i in range(1, 4):
            section = f"Pet {i} Information"
            self.data[section] = {
                "Name": tk.StringVar(),
                "Species": tk.StringVar(),
                "Gender": tk.StringVar(value="N/A"),
                "Breed": tk.StringVar(),
                "Color(s)": tk.StringVar(),
                "Distinguishing characteristic(s)": tk.StringVar(),
                "Stray?": tk.StringVar(value="N/A"),
                "Older than 5 years?": tk.StringVar(value="N/A"),
                "Dog weight range?": tk.StringVar(value="N/A"),
                "Special": tk.StringVar(value="N/A"),
                "Voucher": tk.StringVar(),
                "Sent": tk.StringVar(),
                "Expires": tk.StringVar(),
                "Grant": tk.StringVar()
            }

    def build_frames(self):
        for index, (section, fields) in enumerate(self.data.items()):
            frame = tk.Frame(self, padx=20, pady=20)
            title = tk.Label(frame, text=section, font=('Arial', 16, 'bold'))
            title.grid(row=0, column=0, columnspan=2, pady=(0, 20))

            row = 1
            for label, var in fields.items():
                tk.Label(frame, text=label, anchor='w', width=30).grid(row=row, column=0, sticky='w', padx=5, pady=5)

                if section.startswith("Pet") and label in self.dropdown_fields_common:
                    dropdown = tk.OptionMenu(frame, var, *self.dropdown_fields_common[label])
                    dropdown.config(width=38)
                    dropdown.grid(row=row, column=1, padx=5, pady=5)
                elif section.startswith("Pet") and label in self.gender_field:
                    dropdown = tk.OptionMenu(frame, var, *self.gender_field[label])
                    dropdown.config(width=38)
                    dropdown.grid(row=row, column=1, padx=5, pady=5)
                else:
                    tk.Entry(frame, textvariable=var, width=40).grid(row=row, column=1, padx=5, pady=5)
                row += 1

            # Add "Add another pet?" checkbox if applicable
            if section in ["Pet 1 Information", "Pet 2 Information"]:
                var_flag = tk.IntVar()
                self.add_another_pet_flags[section] = var_flag
                cb = tk.Checkbutton(frame, text="Add another pet?", variable=var_flag)
                cb.grid(row=row, column=0, columnspan=2, pady=(10, 5))
                row += 1

            # Button frame
            button_frame = tk.Frame(frame)
            button_frame.grid(row=row, column=0, columnspan=2, pady=20)

            # Back button (not on first frame)
            if index > 0:
                tk.Button(button_frame, text="Back", command=self.prev_frame).pack(side=tk.LEFT, padx=5)

            # Next or Submit button
            if section in ["Pet 1 Information", "Pet 2 Information"]:
                tk.Button(button_frame, text="Next", command=lambda sec=section: self.handle_pet_next(sec)).pack(side=tk.LEFT, padx=5)
            elif section == "Recipient Information":
                tk.Button(button_frame, text="Next", command=self.next_frame).pack(side=tk.LEFT, padx=5)
            else:
                tk.Button(button_frame, text="Submit", command=self.export_to_excel).pack(side=tk.LEFT, padx=5)

            self.frames.append(frame)

    def handle_pet_next(self, section):
        flag = self.add_another_pet_flags[section].get()
        if section == "Pet 1 Information":
            self.show_frame(2) if flag else self.export_to_excel()
        elif section == "Pet 2 Information":
            self.show_frame(3) if flag else self.export_to_excel()

    def show_frame(self, index):
        for frame in self.frames:
            frame.grid_forget()
        self.frames[index].grid(row=0, column=0, sticky='nsew')
        self.current_frame = index

    def next_frame(self):
        if self.current_frame < len(self.frames) - 1:
            self.show_frame(self.current_frame + 1)

    def prev_frame(self):
        if self.current_frame > 0:
            self.show_frame(self.current_frame - 1)

    def export_to_excel(self):
        # Get last name (assumed 2nd word from "Mail to Name")
        name = self.data["Recipient Information"]["Mail to Name"].get().strip()
        last_name = name.split()[-1] if name else "Unknown"

        # Create filename
        today = datetime.today().strftime('%Y-%m-%d')
        filename = f"{last_name}_{today}.xlsx"
        filepath = os.path.abspath(filename)

        # Write to Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "M-SNAP Form Data"

        row = 1
        for section, fields in self.data.items():
            ws.cell(row=row, column=1, value=section)
            row += 1
            for key, value in fields.items():
                ws.cell(row=row, column=1, value=key)
                ws.cell(row=row, column=2, value=value.get())
                row += 1
            row += 1

        wb.save(filepath)

        # Show success
        messagebox.showinfo("Success", f"Data saved to {filename}")

        # Print the file (Windows only)
        try:
            os.startfile(filepath, "print")
        except Exception as e:
            messagebox.showwarning("Print Error", f"Could not send to printer:\n{e}")

        self.restart_app()

    def restart_app(self):
        for section, fields in self.data.items():
            for key, var in fields.items():
                if section.startswith("Pet") and key in self.dropdown_fields_common:
                    var.set("N/A")
                elif section.startswith("Pet") and key in self.gender_field:
                    var.set("N/A")
                else:
                    var.set("")
        for var in self.add_another_pet_flags.values():
            var.set(0)
        self.show_frame(0)

if __name__ == "__main__":
    app = MSnapApp()
    app.mainloop()
