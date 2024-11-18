from zipfile import ZipFile
from lxml import etree
from sqlalchemy import MetaData, Table, create_engine, select
from sqlalchemy.orm import sessionmaker

# Configuración de la base de datos
DATABASE_URL = "sqlite:////Users/paulalaorga/Development/cartas/Parsed/parsed.db"  # Cambia a tu URL de base de datos
engine = create_engine(DATABASE_URL)
metadata = MetaData()
metadata.reflect(bind=engine)

# Tabla de datos
datos_table = metadata.tables['datos']

# Rutas del archivo
docx_path = '/Users/paulalaorga/Desktop/Cartas/DOC/plantilla.docx'
output_path = '/Users/paulalaorga/Development/cartas/Final/salida_final.docx'


def procesar_documento(docx_path, output_path, replacements):
    with ZipFile(docx_path, 'r') as docx_zip:
        updated_files = {}

        for item in docx_zip.infolist():
            if 'header' in item.filename or 'footer' in item.filename or item.filename == 'word/document.xml':
                xml_content = docx_zip.read(item.filename)
                root = etree.XML(xml_content)

                for paragraph in root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                    full_text = ""
                    runs = paragraph.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
                    for run in runs:
                        if run.text:
                            full_text += run.text

                    for marker, value in replacements.items():
                        if marker in full_text:
                            full_text = full_text.replace(marker, value)

                    for i, run in enumerate(runs):
                        if i == 0:
                            run.text = full_text
                        else:
                            run.text = ""

                    for rPr in paragraph.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr"):
                        highlight = rPr.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight")
                        if highlight is not None:
                            rPr.remove(highlight)

                updated_files[item.filename] = etree.tostring(root, encoding="utf-8", xml_declaration=True)
            else:
                updated_files[item.filename] = docx_zip.read(item.filename)

    with ZipFile(output_path, 'w') as docx_out:
        for filename, content in updated_files.items():
            docx_out.writestr(filename, content)
    print(f"Documento generado: {output_path}")


def main():
    # Crear una sesión
    Session = sessionmaker(bind=engine)
    session = Session()

    # Consultar datos desde la base de datos
    stmt = select(datos_table)
    registro = session.execute(stmt).first()  # Obtiene solo el primer registro

    if registro:
        # Convertir el registro en un diccionario
        registro_dict = dict(registro._mapping)

        # Crear diccionario de reemplazos asegurando que todos los valores sean cadenas
        replacements = {
            "#Agencia#": str(registro_dict.get("Agencia", "Agencia por defecto")),
            "#Telefono#": str(registro_dict.get("Telefono", "123-456-789")),
            "#HorarioAgencia#": str(registro_dict.get("HorarioAgencia", "9:00 AM - 6:00 PM")),
            "#MailCabeceraCapital#": str(registro_dict.get("MailCabeceraCapital", "info@ejemplo.com")),
            "#Interviniente_Nombre#": str(registro_dict.get("Interviniente_Nombre", "Nombre por defecto")),
            "#Direccion_Via#": str(registro_dict.get("Direccion_Via", "Calle Falsa 123")),
            "#Direccion_CodigoPostal#": str(registro_dict.get("Direccion_CodigoPostal", "28000")),
            "#Direccion_Localidad#": str(registro_dict.get("Direccion_Localidad", "Madrid")),
            "#Direccion_DenomProvincia#": str(registro_dict.get("Direccion_DenomProvincia", "Madrid")),
            "#Bar_Code#": str(registro_dict.get("Bar_Code", "123456789")),
            "#dia#": str(registro_dict.get("dia", "01")),
            "#mes#": str(registro_dict.get("mes", "Enero")),
            "#ano#": str(registro_dict.get("ano", "2024")),
            "#Propietario#": str(registro_dict.get("Propietario", "Empresa XYZ")),
            "#Expediente#": str(registro_dict.get("Expediente", "EXP123")),
            "#Deuda#": str(registro_dict.get("Deuda", "5000 €")),
            "#TextoCesion#": str(registro_dict.get("TextoCesion", "Texto legal de cesión")),
            "#Saldo#": str(registro_dict.get("Saldo", "5000 €")),
            "#notariocedente#": str(registro_dict.get("notariocedente", "Notario Ejemplo")),
            "#FechaCesionCedente#": str(registro_dict.get("FechaCesionCedente", "01/01/2023")),
            "#ProtocoloCedente#": str(registro_dict.get("ProtocoloCedente", "PROTO123")),
            "#Exp_Saldo_TEXTO#": str(registro_dict.get("Exp_Saldo_TEXTO", "Cinco mil euros")),
            "#SumaSaldoCapitalAgrupadas#": str(registro_dict.get("SumaSaldoCapitalAgrupadas", "5000")),
            "#CIReferencia#": str(registro_dict.get("CIReferencia", "CI12345")),
            "#cccPropietario#": str(registro_dict.get("cccPropietario", "ES12345678901234567890")),
            "#FechaMasVeintiunDias#": str(registro_dict.get("FechaMasVeintiunDias", "01/02/2024")),
        }

        # Generar archivo para cada registro
        output_file = output_path.replace("salida_final", f"salida_{registro_dict['Codigo']}")
        procesar_documento(docx_path, output_file, replacements)

    session.close()
