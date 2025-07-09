# 📄 PDF to Word Converter with Auto-Cropping

This Python tool converts each page of a multi-page PDF into a high-resolution image, automatically crops the white margins, resizes the image to fit within a Word document's printable area (A4 paper), and saves the result as a `.docx` file — one page per image.

---

## ✨ Features

- ✅ Converts PDF pages to high-resolution images (As required by the Monash University thesis)
- ✂️ Automatically detects and removes white margins
- 🖼 Resizes images to fit inside the page (after accounting for real Word margins)
- 📝 Outputs clean, print-ready Word documents
- ⚙️ Supports adjustable margins, DPI, cropping sensitivity, and shrink factor

---

## 📦 Requirements

- Python 3.10
- [Poppler](https://github.com/oschwartz10612/poppler-windows/releases) (must be installed and added to system PATH)
- Python packages:

```bash
pip install pdf2image python-docx pillow numpy
```

---

## ✅ 🧪 Usage (via Jupyter Notebook)

To use this tool:

1. Open the Jupyter Notebook: `converter.ipynb`
2. Run all cells
3. Change the arguments as needed in the final cell:

```python
auto_crop_and_embed_to_word_fit_page(
    pdf_filename="chapter3.pdf",
    input_dir="inputs",
    output_dir="output",
    dpi=400,
    page_margins=(1.0, 1.0, 1.0, 1.0),  # Top, Bottom, Left, Right in inches
    shrink_scale=0.90,                 # Shrink image after fitting
    threshold=245                      # Brightness threshold for cropping
)
```

> 📂 Place your PDFs inside the `inputs/` folder.  
> 📄 Output `.docx` files will appear in the `output/` folder.

---

## ⚙️ Parameters

| Parameter       | Description                                                           | Default            |
|----------------|-----------------------------------------------------------------------|--------------------|
| `pdf_filename`  | Name of the PDF file in the `inputs/`                               | —                  |
| `input_dir`     | Directory containing the PDF                                           | `"inputs"`         |
| `output_dir`    | Directory to save the generated Word document                         | `"output"`         |
| `dpi`           | Resolution for converting PDF pages to images                         | `400`              |
| `page_margins`  | Margins for the Word document (top, bottom, left, right), in inches   | `(1.0, 1.0, 1.0, 1.0)` |
| `shrink_scale`  | Additional downscaling after fitting image to page                    | `0.90`             |
| `threshold`     | Brightness threshold for detecting white margins in images (0–255)    | `245`              |

---

## 🧾 Output

After running the function, you'll find the output in:

```
output/your_file.docx
```

Each page in the `.docx` file corresponds to one cropped and resized PDF page.

---

## 📝 Notes

- This tool assumes A4 page size (8.27 × 11.69 inches)
- Actual Word margins are applied — so output is WYSIWYG when printing
- Images are inserted using `docx` with exact dimensions (width × height)
- Temporary image files are stored in `temp_img_autocrop/` and removed after completion

---

## 📚 License

This project is licensed under the MIT License.

---

## 👤 Author

**Viet Long Bui**  
📧 [viet.bui1@monash.edu](mailto:viet.bui1@monash.edu)

---

## 🌐 Related Tools

- [`pdf2image`](https://github.com/Belval/pdf2image) — Convert PDFs to images
- [`python-docx`](https://python-docx.readthedocs.io/) — Create and edit Word documents
- [`poppler-utils`](https://poppler.freedesktop.org/) — PDF rendering backend required by `pdf2image`
