import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image
from barcode import EAN13
from barcode.writer import ImageWriter
from io import BytesIO

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Etiquetas")

        self.file_path = None
        self.load_data()

        self.ref_var = tk.StringVar()
        self.ref_var.trace_add("write", self.update_table)

        self.create_widgets()

    def load_data(self):
        if self.file_path:
            try:
                self.df = pd.read_excel(self.file_path)
            except pd.errors.EmptyDataError:
                self.df = pd.DataFrame(columns=["ref", "produtoDesc", "codigoBarras"])
        else:
            self.df = pd.DataFrame(columns=["ref", "produtoDesc", "codigoBarras"])

    def create_widgets(self):
        # Frame superior
        self.frame_top = ttk.Frame(self.root)
        self.frame_top.pack(pady=10)

        self.tree = ttk.Treeview(self.frame_top, columns=("Ref", "Produto", "Código de Barras"), show="headings")
        self.tree.heading("Ref", text="Ref")
        self.tree.heading("Produto", text="Produto")
        self.tree.heading("Código de Barras", text="Código de Barras")
        self.tree.pack(side="left")

        self.scrollbar = ttk.Scrollbar(self.frame_top, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.scrollbar.set)

        # Frame inferior
        self.frame_bottom = ttk.Frame(self.root)
        self.frame_bottom.pack(pady=10)

        self.ref_label = ttk.Label(self.frame_bottom, text="Selecione a REF:")
        self.ref_label.grid(row=0, column=0, padx=5, pady=5)

        self.ref_combobox = ttk.Combobox(self.frame_bottom, textvariable=self.ref_var, values=self.df["ref"].unique())
        self.ref_combobox.grid(row=0, column=1, padx=5, pady=5)

        self.load_file_button = ttk.Button(self.frame_bottom, text="Carregar Arquivo", command=self.load_file)
        self.load_file_button.grid(row=0, column=2, padx=5, pady=5)

        self.generate_pdf_button = ttk.Button(self.frame_bottom, text="Gerar PDF", command=self.generate_pdf)
        self.generate_pdf_button.grid(row=0, column=3, padx=5, pady=5)
        
    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if file_path:
            self.file_path = file_path
            self.load_data()
            self.update_table()

    def update_table(self, *args):
        selected_ref = self.ref_var.get()
        if selected_ref:
            selected_data = self.df[self.df["ref"] == selected_ref]
        else:
            selected_data = self.df

        self.tree.delete(*self.tree.get_children())

        for index, row in selected_data.iterrows():
            self.tree.insert("", "end", values=(row["ref"], row["produtoDesc"], row["codigoBarras"]))

    def generate_barcode_image(self, barcode_data):
        barcode = EAN13(barcode_data, writer=ImageWriter())
        buffer = BytesIO()
        barcode.write(buffer)
        return Image(buffer, width=100, height=30)

    def generate_pdf(self):
            selected_ref = self.ref_var.get()

            if not self.df.empty:
                if selected_ref:
                    selected_data = self.df[self.df["ref"] == selected_ref]
                else:
                    selected_data = self.df
            else:
                messagebox.showinfo("Aviso", "O DataFrame está vazio. Carregue dados antes de gerar o PDF.")
                return

            buffer = BytesIO()
            pdf = SimpleDocTemplate(buffer, pagesize=letter)

            data = [["Ref", "Produto", "Código de Barras"]]

            for index, row in selected_data.iterrows():
                ref = str(row['ref'])
                produto_desc = str(row['produtoDesc'])
                codigo_barras = str(row['codigoBarras'])

                # Gera o código de barras
                barcode_image = self.generate_barcode_image(codigo_barras)

                data.append([ref, produto_desc, barcode_image]) 

            # Adiciona a tabela ao PDF
            table_data = []
            for row in data:
                table_data.append(row)

            table = Table(table_data)

            # Adiciona estilo à tabela
            style = TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                                ('GRID', (0, 0), (-1, -1), 1, colors.black)])

            table.setStyle(style)

            # Adiciona as imagens dos códigos de barras à lista de elementos
            elements = [table]

            pdf.build(elements)

            # Salva o PDF
            pdf_output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Arquivos PDF", "*.pdf")])
            if pdf_output_path:
                with open(pdf_output_path, "wb") as pdf_output_file:
                    pdf_output_file.write(buffer.getvalue())
                messagebox.showinfo("PDF Gerado", f"O PDF foi gerado com sucesso em {pdf_output_path}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()