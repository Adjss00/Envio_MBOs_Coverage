import pandas as pd

class ExcelReader:
    def __init__(self, excel_file_path, sheet_names):
        self.excel_file_path = excel_file_path
        self.sheet_names = sheet_names

    def read_excel_sheets(self):
        try:
            # Cargar el archivo Excel
            excel_data = pd.ExcelFile(self.excel_file_path)

            # Inicializar un diccionario para almacenar los DataFrames
            dataframes = {}

            # Iterar sobre las hojas especificadas
            for sheet_name in self.sheet_names:
                # Leer la hoja y almacenarla en el diccionario
                df = excel_data.parse(sheet_name)
                dataframes[sheet_name] = df

            # Aplicar la lógica de conversión a cada DataFrame
            for sheet_name, df in dataframes.items():
                self.apply_conversion_logic(df)

                print(f"DataFrame para la hoja '{sheet_name}' después de aplicar la lógica de conversión:")
                print(df)
                print("\n" + "-"*40 + "\n")

            return dataframes

        except Exception as e:
            print(f"Error al leer el archivo Excel: {e}")

    def apply_conversion_logic(self, df):
        for index, row in df.iterrows():
            mbo_value = row.get("MBO")

            if mbo_value == "Seguros" or mbo_value == "Coverage":
                df.at[index, "Target"] = f"{int(row['Target'] * 100)}%"
                df.at[index, "Actual"] = f"{int(row['Actual'] * 100)}%"
            elif mbo_value == "Factoraje" or mbo_value == "Fleet" or mbo_value == "Creación Pipeline":
                df.at[index, "Target"] = f"${row['Target']}"
                df.at[index, "Actual"] = f"${row['Actual']}"
            elif mbo_value == "Pipeline a final del 4Q":
                df.at[index, "Target"] = f"${row['Target']}"
                df.at[index, "Actual"] = f"${row['Actual']}"

            # Update "% Cumplimiento" to percentage, handling missing or non-numeric values
            try:
                cumple_value = float(row.get('% Cumplimiento ', 0))
                df.at[index, "% Cumplimiento "] = f"{int(cumple_value * 100)}%"
            except (ValueError, TypeError):
                print("Advertencia: El campo '% Cumplimiento' no es un valor numérico.")
