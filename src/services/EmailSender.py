import os
import pandas as pd
import win32com.client

class EmailSender:
    def __init__(self, directory_path, originadores):
        self.directory_path = directory_path
        self.originadores = originadores

    def compare_and_send_email(self, excel_file_path):
        try:
            # Obtener el nombre del archivo sin extensión
            filename_no_extension = os.path.splitext(os.path.basename(excel_file_path))[0]

            # Verificar coincidencias con los originadores
            matching_originadores = {key: value for key, value in self.originadores.items() if key.lower() in filename_no_extension.lower()}

            if matching_originadores:
                # Obtener el nombre del originador sin guion bajo
                originador_name = list(matching_originadores.keys())[0].replace("_", " ")

                # Verificar si la cadena "MBOs_" está presente en el nombre del archivo
                if "MBOs_" in filename_no_extension:
                    # Obtener la parte del nombre del archivo después de "MBOs_"
                    file_suffix = filename_no_extension.split("MBOs_")[1]

                    # Verificar coincidencias exactas con los originadores
                    exact_matching_originadores = {key: value for key, value in matching_originadores.items() if key.lower() == file_suffix.lower()}

                    if exact_matching_originadores:
                        # Crear el cuerpo del correo electrónico con el nuevo texto y una tabla HTML
                        email_body = f"<html><body>"
                        email_body += f"<p style='color: #1A5276;'>Hola <strong>{originador_name}</strong>,</p>"
                        email_body += "<p style='color: #1A5276;'>Te comparto tus MBOs, te recordamos que:</p>"
                        email_body += "<ul>"
                        email_body += "<li style='color: #1A5276;'>Los MBOs de <strong>'Penetración de Productos'</strong> y <strong>'Clientes nuevos'</strong> se contabilizan anualmente (si aplica).</li>"
                        email_body += "<li style='color: #1A5276;'>Para los MBOs de <strong>'Coverage'</strong> tienes hasta el 5 de Enero de 2024 para completarlo (recuerda que el registro debe estar fechado de Octubre-Diciembre).</li>"
                        email_body += "<li style='color: #1A5276;'>Para el resto de los MBOs tienes hasta el 31 de diciembre 2023.</li>"
                        email_body += "</ul>"

                        # Leer el CSV y agregar los datos a la tabla HTML
                        df = pd.read_csv(excel_file_path)

                        # Reemplazar 'nan' con 0 en toda la DataFrame
                        df = df.fillna(0)

                        # Excluir columnas específicas
                        excluded_columns = ['Column1', 'Nombre', 'Apellido', 'Puesto', 'ID', "Notas", "meta", "monto seguro", "monto fondeado"]
                        df = df.drop(excluded_columns, axis=1, errors='ignore')

                        email_body += "<table border='1' style='border-collapse: collapse; border-color: #1A5276;'>"
                        email_body += "<tr style='background-color: #1A5276; color: #FFFFFF; font-weight: bold; border-color: #1A5276;'>"
                        email_body += "".join([f"<th>{col}</th>" for col in df.columns]) + "</tr>"
                        for _, row in df.iterrows():
                            email_body += "<tr>"
                            for col, value in row.items():
                                # Aplicar color a las celdas con valores de porcentaje
                                if '%' in str(value):
                                    percentage = float(str(value).replace('%', ''))
                                    color = self.interpolate_color("#F1948A", "#52BE80", percentage)
                                    email_body += f"<td style='background-color: {color}; color: #1A5276; border-color: #1A5276;'>{value}</td>"
                                else:
                                    email_body += f"<td style='background-color: #D4E6F1; color: #1A5276; border-color: #1A5276;'>{value}</td>"
                            email_body += "</tr>"

                        email_body += "</table>"
                        email_body += "</body></html>"

                        # Enviar correo electrónico a través de Outlook
                        outlook = win32com.client.Dispatch('Outlook.Application')
                        mail = outlook.CreateItem(0)
                        mail.Subject = "Coincidencia de Originador"

                        # Añadir saludo de despedida
                        email_body += "<p style='color: #1A5276;'>Saludos cordiales, <strong>Jesús Sanchez</strong></p>"

                        mail.HTMLBody = email_body
                        mail.To = ", ".join(exact_matching_originadores.values())
                        mail.Send()

                        print(f"Correo enviado para coincidencia en {file_suffix}")

        except Exception as e:
            print(f"Error en el procesamiento: {e}")

    def process_files(self):
        try:
            # Iterar sobre todos los archivos en el directorio
            for filename in os.listdir(self.directory_path):
                if filename.endswith(".csv"):
                    file_path = os.path.join(self.directory_path, filename)
                    self.compare_and_send_email(file_path)
        except Exception as e:
            print(f"Error en el bucle principal: {e}")

    @staticmethod
    def interpolate_color(color1, color2, percentage):
        """
        Interpola entre dos colores en función del porcentaje.
        :param color1: Color inicial en formato hexadecimal.
        :param color2: Color final en formato hexadecimal.
        :param percentage: Porcentaje entre 0 y 100.
        :return: Color resultante en formato hexadecimal.
        """
        color1 = [int(color1[i:i + 2], 16) for i in (1, 3, 5)]
        color2 = [int(color2[i:i + 2], 16) for i in (1, 3, 5)]

        result_color = [
            round(color1[i] + (color2[i] - color1[i]) * percentage / 100)
            for i in range(3)
        ]

        return "#{:02X}{:02X}{:02X}".format(*result_color)