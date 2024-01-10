
from services.DataExporter import DataExporter, project_dir
from services.EmailSender import EmailSender
from services.ExcelReader import ExcelReader

if __name__ == "__main__":
    excel_info = {"excel_file_path": "out/update/Mbos_4Q_2023.xlsx", "sheet_names": ["MBOs"]}
    excel_reader = ExcelReader(excel_info["excel_file_path"], excel_info["sheet_names"])
    dataframes = excel_reader.read_excel_sheets()

    # Crear una instancia de DataExporter
    exporter = DataExporter(project_dir, "out/archive")

    # Iterar sobre los DataFrames y exportar a CSV
    for sheet_name, df in dataframes.items():
        exporter.export_df_to_csv(df, sheet_name)

# Directorio que contiene los archivos csv
directory_path = "out/archive"

# Lista de originadores
originadores = {
    "Janira_Gonzalez": "janira.gonzalez@engen.com.mx",
    "Guadalupe_Villegas": "guadalupe.villegas@engen.com.mx"
    # ... (otros originadores)
}

# Crear instancia de la clase EmailSender y procesar archivos
email_sender = EmailSender(directory_path, originadores)
email_sender.process_files()