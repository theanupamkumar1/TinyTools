import tkinter as tk
from tkinter import filedialog
import glob
import comtypes.client


def ppt_to_pdf(ppt_file, pdf_file):
  """Converts a PowerPoint file (ppt or pptx) to PDF.

  Args:
    ppt_file: Path to the input PowerPoint file.
    pdf_file: Path to the output PDF file.
  """
  powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
  powerpoint.Visible = False

  try:
    deck = powerpoint.Presentations.Open(ppt_file)
    deck.SaveAs(pdf_file, 32)  # 32 for PDF format
    deck.Close()
  except Exception as e:
    print(f"Error converting {ppt_file}: {e}")
  finally:
    powerpoint.Quit()


def select_files():
  """Opens a file dialog to select PPT files."""
  filenames = filedialog.askopenfilenames(title="Select PPT Files", filetypes=[("PowerPoint Presentations", "*.pptx *.ppt")])
  for filename in filenames:
    pdf_file = filename.replace(".pptx", ".pdf") if filename.endswith(".pptx") else filename.replace(".ppt", ".pdf")
    ppt_to_pdf(filename, pdf_file)
  print("Conversion completed!")


root = tk.Tk()
root.title("PPT to PDF Converter")

button = tk.Button(root, text="Select Files", command=select_files)
button.pack(padx=10, pady=10)

root.mainloop()
