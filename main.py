from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import Tk, Canvas, PhotoImage
import subprocess
from calculation import create_source_sheet, create_file, delete_default_sheet


def relative_to_assets(path: str) -> Path:
    current_path = Path(__file__).parent.absolute()  # Get the current directory
    assets_path = current_path / Path("assets/frame0")  # Set assets path
    return assets_path / Path(path)


# Global variables
file_path = None
data_frame = None

window = Tk()
window.title("Метод ПУЗ-ОРХ")
window.geometry("718x758")
window.configure(bg="#06010F")


def select_excel_file(event):
    # Open file dialog to select an Excel file
    global file_path

    default_file_name = 'table'

    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=default_file_name,
                                             filetypes=[("Excel Files", "*.xlsx")])

    create_file(file_path)

    create_source_sheet(file_path, 1, 5, True)

    delete_default_sheet(file_path)


canvas = Canvas(window, bg="#06010F", height=758, width=718, bd=0, highlightthickness=0, relief="ridge")

canvas.place(x=0, y=0)
canvas.create_text(180.0, 115.0, anchor="nw", text="Метод ПУЗ-ОРХ", fill="#FFFFFF",
                   font=("RalewayRoman SemiBold", 48 * -1))

image_image_1 = PhotoImage(file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(359.0, 79.0, image=image_image_1)

image_image_2 = PhotoImage(file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(359.0, 379.0, image=image_image_2, anchor="center")


def nextCall():
    if file_path:
        subprocess.call(["python", "step.py", "--file", file_path])
    else:
        messagebox.showinfo("Error", "Сначала создайте и сохраните файл.")


canvas.create_text(298.0, 569.0, anchor="nw", text="Загрузить", fill="#FFFFFF",
                   font=("PlusJakartaSansRoman ExtraBold", 24 * -1))

image_image_3 = PhotoImage(file=relative_to_assets("image_3.png"))
image_3 = canvas.create_image(358.0, 379.0, image=image_image_3)

button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
image_4 = canvas.create_image(580.0, 700.0, image=button_image_1)

canvas.tag_bind(image_4, "<Button-1>", lambda event: nextCall())
canvas.tag_bind(image_3, "<Button-1>", select_excel_file)
canvas.create_text(120.0, 614.0, anchor="nw", text="Пожалуйста, выберите где вы хотите сохранить файл", fill="#82819D",
                   font=("Inter Medium", 20 * -1))

window.resizable(False, False)
window.mainloop()
