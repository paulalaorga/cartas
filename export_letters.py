from zipfile import ZipFile
from lxml import etree
from sqlalchemy import create_engine, MetaData, select
from sqlalchemy.orm import sessionmaker

# Configuración de la base de datos
DATABASE_URL = "sqlite:////Users/paulalaorga/Development/cartas/Parsed/parsed.db"
engine = create_engine(DATABASE_URL)
metadata = MetaData()
metadata.reflect(bind=engine)

# Tabla de datos
datos_table = metadata.tables['datos']

# Rutas del archivo
docx_path = '/Users/paulalaorga/Development/cartas/Raw Data/Plantilla.docx'
output_path = '/Users/paulalaorga/Development/cartas/Final/salida_final.docx'


def procesar_documento(docx_path, output_path, replacements):
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    with ZipFile(docx_path, 'r') as docx_zip:
        updated_files = {}

        for item in docx_zip.infolist():
            try:
                # Procesar solo archivos XML relevantes
                if item.filename.endswith(".xml") and ('word/' in item.filename or 'header' in item.filename or 'footer' in item.filename):
                    xml_content = docx_zip.read(item.filename)
                    root = etree.XML(xml_content)

                     # Procesar todos los nodos de propiedades de estilo y eliminar resaltados
                    for run_properties in root.findall(".//w:rPr", namespaces):
                        highlight = run_properties.find("w:highlight", namespaces)
                        if highlight is not None:
                            run_properties.remove(highlight)

                    # Procesar todos los elementos relevantes en el documento
                    for element in root.findall(".//*", namespaces):
                        # Procesar párrafos
                        if element.tag.endswith("p"):  # <w:p>
                            runs = element.findall(".//w:r", namespaces)
                            text_nodes = [run.find(".//w:t", namespaces) for run in runs if run.find(".//w:t", namespaces) is not None]

                            # Reconstruir y reemplazar texto
                            full_text = "".join(node.text or "" for node in text_nodes)
                            for marker, value in replacements.items():
                                if marker in full_text:
                                    full_text = full_text.replace(marker, value)

                            # Redistribuir el texto preservando estilo
                            remaining_text = full_text
                            for i, (run, node) in enumerate(zip(runs, text_nodes)):
                                    run_properties = run.find("w:rPr", namespaces)

                                    if i == 0:
                                        node.text = remaining_text
                                        if run_properties is not None:
                                           run.insert(0, run_properties)

                                    else:
                                      node.text = ""  # Limpiar nodos adicionales
                                

                        # Procesar celdas de tablas
                        elif element.tag.endswith("tc"):  # <w:tc>
                            text_nodes = element.findall(".//w:t", namespaces)
                            full_text = "".join(node.text or "" for node in text_nodes)
                            for marker, value in replacements.items():
                                if marker in full_text:
                                    full_text = full_text.replace(marker, value)

                            # Redistribuir el texto en las celdas
                            remaining_text = full_text
                            for node in text_nodes:
                                if remaining_text:
                                    node.text = remaining_text
                                    remaining_text = ""
                                else:
                                    node.text = ""

                        # Procesar encabezados y pies de página
                        elif element.tag.endswith("hdr") or element.tag.endswith("ftr"):  # <w:hdr>, <w:ftr>
                            text_nodes = element.findall(".//w:t", namespaces)
                            full_text = "".join(node.text or "" for node in text_nodes)
                            for marker, value in replacements.items():
                                if marker in full_text:
                                    full_text = full_text.replace(marker, value)

                            # Redistribuir el texto
                            remaining_text = full_text
                            for node in text_nodes:
                                if remaining_text:
                                    node.text = remaining_text
                                    remaining_text = ""
                                else:
                                    node.text = ""

                    # Guardar los cambios en el archivo
                    updated_files[item.filename] = etree.tostring(root, encoding="utf-8", xml_declaration=True)
                else:
                    # Copiar archivos no relevantes sin modificar
                    updated_files[item.filename] = docx_zip.read(item.filename)

            except Exception as e:
                print(f"Error procesando {item.filename}: {e}")

    # Escribir el archivo modificado
    with ZipFile(output_path, 'w') as docx_out:
        for filename, content in updated_files.items():
            docx_out.writestr(filename, content)


    print(f"Documento generado: {output_path}")


    # Escribir el archivo modificado
    with ZipFile(output_path, 'w') as docx_out:
        for filename, content in updated_files.items():
            docx_out.writestr(filename, content)

    print(f"Documento generado: {output_path}")


    # Escribir el archivo modificado
    with ZipFile(output_path, 'w') as docx_out:
        for filename, content in updated_files.items():
            docx_out.writestr(filename, content)

    print(f"Documento generado: {output_path}")




def main():
    # Crear una sesión de base de datos
    Session = sessionmaker(bind=engine)
    session = Session()

    # Consultar datos de la base de datos
    stmt = select(datos_table)
    registro = session.execute(stmt).first()

    if registro:
        # Convertir el registro en un diccionario
        registro_dict = dict(registro._mapping)


        replacements = {
            "#Agencia#": str(registro_dict.get("agencia")),
            "#Telefono#": str(registro_dict.get("Telefono")),
            "#HorarioAgencia#": str(registro_dict.get("HorarioAgencia")),
            "#MailCabeceraCapital#": str(registro_dict.get("MailCabeceraCapital")),
            "#Interviniente_Nombre#": str(registro_dict.get("Interviniente_Nombre")),
            "#Direccion_Via#": str(registro_dict.get("Direccion_Via")),
            "#Direccion_CodigoPostal#": str(registro_dict.get("Direccion_CodigoPostal")),
            "#Direccion_Localidad#": str(registro_dict.get("Direccion_Localidad")),
            "#Direccion_DenomProvincia#": str(registro_dict.get("Direccion_DenomProvincia")),
            "#Bar_Code#": str(registro_dict.get("Bar_Code")),
            "#dia#": str(registro_dict.get("dia")),
            "#mes#": str(registro_dict.get("mes")),
            "#ano#": str(registro_dict.get("ano")),
            "#Propietario#": str(registro_dict.get("Propietario")),
            "#Expediente#": str(registro_dict.get("Expediente")),
            "#Deuda#": str(registro_dict.get("Deuda")),
            "#TextoCesion#": str(registro_dict.get("TextoCesion")),
            "#Saldo#": str(registro_dict.get("Saldo")),
            "#notariocedente#": str(registro_dict.get("notariocedente")),
            "#FechaCesionCedente#": str(registro_dict.get("FechaCesionCedente")),
            "#ProtocoloCedente#": str(registro_dict.get("ProtocoloCedente")),
            "#Exp_Saldo_TEXTO#": str(registro_dict.get("Exp_Saldo_TEXTO")),
            "#SumaSaldoCapitalAgrupadas#": str(registro_dict.get("SumaSaldoCapitalAgrupadas")),
            "#CIReferencia#": str(registro_dict.get("CIReferencia")),
            "#cccPropietario#": str(registro_dict.get("cccPropietario")),
            "#FechaMasVeintiunDias#": str(registro_dict.get("FechaMasVeintiunDias")),
        }

        procesar_documento(docx_path, output_path, replacements)

    session.close()


if __name__ == "__main__":
    main()
