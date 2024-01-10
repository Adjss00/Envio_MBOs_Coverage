import os
import pandas as pd

# Obtén la ruta al directorio principal del proyecto
current_dir = os.path.dirname(os.path.realpath(__file__))
project_dir = os.path.join(current_dir, )

class DataExporter:
    def __init__(self, project_dir, output_dir):
        self.project_dir = project_dir
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)

    def export_df_to_csv(self, df, file_name_prefix):
        try:
            # Group the DataFrame by the "Originador" column
            grouped_df = df.groupby("Originador")

            # Iterate over groups and export each to a separate CSV file
            for originador, originador_df in grouped_df:
                output_file_path = os.path.join(self.output_dir, f'{file_name_prefix}_{originador.replace(" ", "_")}.csv')
                originador_df.to_csv(output_file_path, index=False)
                print(f"Datos exportados a {output_file_path}")

        except Exception as e:
            print(f"Error al exportar datos a CSV: {e}")