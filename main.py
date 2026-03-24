import tkinter as tk
from tkinter import filedialog, messagebox
import os

from data_processor import DataProcessor
from kpi_matcher import KPIMatcher
from report_generator import ReportGenerator

class KPIGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Générateur de Rapport KPI 3G")
        self.root.geometry("600x300")

        self.template_path = tk.StringVar()
        self.raw_data_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        # Configuration des polices
        font_label = ("Arial", 10, "bold")
        font_entry = ("Arial", 10)

        # Ligne Template
        tk.Label(self.root, text="Fichier Template:", font=font_label).grid(row=0, column=0, padx=10, pady=15, sticky="e")
        tk.Entry(self.root, textvariable=self.template_path, width=40, font=font_entry).grid(row=0, column=1, padx=10)
        tk.Button(self.root, text="Parcourir", command=self.browse_template).grid(row=0, column=2, padx=10)

        # Ligne Raw Data
        tk.Label(self.root, text="Données Brutes:", font=font_label).grid(row=1, column=0, padx=10, pady=15, sticky="e")
        tk.Entry(self.root, textvariable=self.raw_data_path, width=40, font=font_entry).grid(row=1, column=1, padx=10)
        tk.Button(self.root, text="Parcourir", command=self.browse_raw).grid(row=1, column=2, padx=10)

        # Ligne Destination
        tk.Label(self.root, text="Dossier Sortie:", font=font_label).grid(row=2, column=0, padx=10, pady=15, sticky="e")
        tk.Entry(self.root, textvariable=self.output_path, width=40, font=font_entry).grid(row=2, column=1, padx=10)
        tk.Button(self.root, text="Parcourir", command=self.browse_output).grid(row=2, column=2, padx=10)

        # Bouton Générer
        self.gen_btn = tk.Button(self.root, text="Générer le Rapport", font=("Arial", 12, "bold"), bg="green", fg="white", command=self.generate)
        self.gen_btn.grid(row=3, column=0, columnspan=3, pady=30)

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.template_path.set(path)

    def browse_raw(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.raw_data_path.set(path)

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def generate(self):
        template = self.template_path.get()
        raw = self.raw_data_path.get()
        out = self.output_path.get()

        if not template or not raw or not out:
            messagebox.showerror("Erreur", "Veuillez sélectionner tous les fichiers nécessaires.")
            return

        self.gen_btn.config(text="Génération en cours...", state=tk.DISABLED)
        self.root.update()

        try:
            # Step 1: Process data
            processor = DataProcessor(raw)
            df = processor.filter_last_10_hours()

            # Step 2: Match KPIs
            matcher = KPIMatcher(df)
            shaped, times, kpis = matcher.process()

            # Step 3: Generate report
            generator = ReportGenerator(template, out, shaped, times, kpis)
            generator.generate()

            messagebox.showinfo("Succès", f"Le rapport a été généré avec succès :\n{out}")

        except Exception as e:
            messagebox.showerror("Erreur de génération", f"Une erreur est survenue :\n{str(e)}")

        finally:
            self.gen_btn.config(text="Générer le Rapport", state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = KPIGeneratorApp(root)
    root.mainloop()
