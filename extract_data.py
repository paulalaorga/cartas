import csv
import re
from datetime import datetime

# Definir los nombres de las columnas
column_names = [
    "Bar_Code", "Tipo_Carta", "Codigo",
    "Interviniente_Nombre", "Direccion_Via", "Direccion_Localidad", "Telefono", "HorarioAgencia", "MailCabeceraCapital", 
    "Direccion_DenomProvincia", "Direccion_CodigoPostal", "codVerif", "dia", "mes", "ano","Interviniente_Nombre2", "Propietario", "cccPropietario", "FechaMasVeintiunDias", "Expediente", "Deuda",
    "TextoCesion", "Saldo", "notariocedente", "FechaCesionCedente", "ProtocoloCedente", "Exp_Saldo_TEXTO", "SumaSaldoCapitalAgrupadas", "CIReferencia"
]

    # Función para transformar la fecha
def transformar_fecha(fecha_original):
        try:
         # Parsear la fecha original al formato datetime
         fecha_formateada = datetime.strptime(fecha_original, "%b %d %Y %I:%M%p")
        # Convertir la fecha al formato deseado
         return fecha_formateada.strftime("%d/%m/%Y")
        except ValueError:
            # En caso de error, retornar la fecha original o un valor vacío
            return ""

def process_line(line, line_number):
    # Elimina espacios en blanco extra y divide en datos básicos y contacto
    partes = re.split(r'\s{2,}', line.strip(), maxsplit=6)
      # Validar si partes tiene al menos 6 elementos
    if len(partes) < 6:
            print(f"Línea ignorada por datos incompletos: {line}")
            return None
    

    # Extrae campos básicos de la parte inicial
    datos_basicos = partes[:6]  # Los primeros 6 elementos son datos básicos
    datos_contacto = partes[6].split('#') if len(partes) > 6 and '#' in partes[6] else []

    # Extraer los dos primeros dígitos y los siguientes cinco para código postal
    direccion_denom_provincia = datos_contacto[0][:2] if len(datos_contacto) > 0 and len(datos_contacto[0]) >= 2 else ""
    direccion_codigo_postal = datos_contacto[0][2:7] if len(datos_contacto) > 0 and len(datos_contacto[0]) >= 7 else ""
    cod_verif = datos_contacto[0][7:37] if len(datos_contacto) > 0 and len(datos_contacto[0]) > 37 else ""

     # Transformar el Bar_Code al formato correcto
    bar_code_original = datos_basicos[0].strip() if len(datos_basicos) > 0 else ""
    if line_number == 1:
            # Cortar 3 caracteres para la primera línea
            bar_code = f"{bar_code_original[3:]}" if len(bar_code_original) > 3 else bar_code_original
    else:
            # Cortar 2 caracteres para el resto de las líneas
            bar_code = f"{bar_code_original[2:]}" if len(bar_code_original) > 2 else bar_code_original
    # Buscar cuenta bancaria en datos_contacto
    cuenta_bancaria = ""
    for dato in datos_contacto:
        if re.match(r"^ES\d{2}", dato.strip()):
            cuenta_bancaria = dato.strip()
            break

    # Búsqueda de la fecha de vencimiento en el texto completo de la línea
    fecha_vencimiento = ""
    fecha_vencimiento_match = re.search(r"\b\d{2}/\d{2}/\d{4}\b", line)
    if fecha_vencimiento_match:
        fecha_vencimiento = fecha_vencimiento_match.group(0)
    
    # Transformar FechaCesionCedente
    fecha_cesion_cedente = transformar_fecha(datos_contacto[14].strip()) if len(datos_contacto) > 14 else ""
    


    # Crear una lista con los valores procesados en el orden de las columnas
    documento = [
        bar_code,
        datos_basicos[1].strip(),  # Tipo_Carta
        datos_basicos[2].strip(),  # Codigo
        datos_basicos[3].strip(),  # Interviniente_Nombre
        datos_basicos[4].strip(),  # Direccion_Via
        datos_basicos[5].strip(),  # Direccion_Localidad
        datos_contacto[1].strip() if len(datos_contacto) > 1 else "",  # Telefono
        datos_contacto[2].strip() if len(datos_contacto) > 2 else "",  # HorarioAgencia
        datos_contacto[3].strip() if len(datos_contacto) > 3 else "",  # MailCabeceraCapital
        direccion_denom_provincia,
        direccion_codigo_postal,
        cod_verif,
        datos_contacto[4].strip() if len(datos_contacto) > 4 else "",  # dia 
        datos_contacto[5].strip() if len(datos_contacto) > 5 else "",  # mes 
        datos_contacto[6].strip() if len(datos_contacto) > 6 else "",  # ano 
        datos_contacto[7].strip() if len(datos_contacto) > 7 else "",  # Interviniente2 
        datos_contacto[8].strip() if len(datos_contacto) > 8 else "",  # Propietario
        cuenta_bancaria,  # cccPropietario
        fecha_vencimiento,  # FechaMasVeintiunDias
        datos_contacto[9].strip() if len(datos_contacto) > 9 else "",  # Expediente
        datos_contacto[10].strip() if len(datos_contacto) > 10 else "",  # Deuda
        datos_contacto[11].strip() if len(datos_contacto) > 11 else "",  # TextoCesion
        datos_contacto[12].strip() if len(datos_contacto) > 12 else "", # Saldo
        datos_contacto[13].strip() if len(datos_contacto) > 13 else "", # notariocedente
        fecha_cesion_cedente, # FechaCesionCedente
        datos_contacto[15].strip() if len(datos_contacto) > 15 else "", # ProtocoloCedente
        datos_contacto[17].strip() if len(datos_contacto) > 17 else "", # Exp_Saldo_TEXTO
        datos_contacto[18].strip() if len(datos_contacto) > 18 else "", # SumaSaldoCapitalAgrupadas
        datos_contacto[22].strip() if len(datos_contacto) > 22 else ""  # CIReferencia
 
    ]

    # Retornar los campos procesados solo si al menos el 50% de los campos contienen datos
    if sum(bool(campo) for campo in documento) >= len(documento) // 2:
        return documento
    else:
        return None

def convert_file(input_filename, output_filename):
    try:
        with open(input_filename, 'r', encoding='utf-8') as infile, \
             open(output_filename, 'w', encoding='utf-8', newline='') as outfile:
            
            csv_writer = csv.writer(outfile, delimiter='|')
            csv_writer.writerow(column_names)  # Escribir nombres de columnas
            
            for line_number, line in enumerate(infile, start=1):
                if line.strip():  # Ignorar líneas vacías
                    processed_parts = process_line(line, line_number)
                    if processed_parts:
                        csv_writer.writerow(processed_parts)
            print(f"Archivo '{output_filename}' creado exitosamente.")
    except FileNotFoundError:
        print(f"Error: El archivo '{input_filename}' no se encontró.")
    except IOError as e:
        print(f"Error de entrada/salida: {e}")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

# Convertir archivo
if __name__ == "__main__":
    input_filename = "/Users/paulalaorga/Development/cartas/Raw Data/input.txt"
    output_filename = "/Users/paulalaorga/Development/cartas/Raw Data/output.csv"
    convert_file(input_filename, output_filename)
