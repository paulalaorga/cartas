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

def process_document(docx_path, output_path, replacements):
    namespaces = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',  # Main WordprocessingML namespace
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',  # Wordprocessing Drawing namespace
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',  # DrawingML namespace
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',  # Relationships namespace
    'v': 'urn:schemas-microsoft-com:vml',  # VML namespace
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',  # MathML namespace
    'w10': 'urn:schemas-microsoft-com:office:word',  # Office Word namespace
    'sl': 'http://schemas.openxmlformats.org/schemaLibrary/2006/main',  # Schema Library namespace
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml'  # Word 2006 namespace
}
    
    with ZipFile(docx_path, 'r') as docx:
        modified_files = {}
        
        # Process all document parts
        for item in docx.filelist:
            if item.filename.endswith('.xml') and ('word/' in item.filename):
                content = docx.read(item.filename)
                tree = etree.fromstring(content)
                process_all_elements(tree, replacements, namespaces)
                modified_files[item.filename] = etree.tostring(tree, xml_declaration=True, encoding='UTF-8')
        
        # Write modified document
        with ZipFile(output_path, 'w') as output:
            for item in docx.filelist:
                if item.filename in modified_files:
                    output.writestr(item, modified_files[item.filename])
                else:
                    output.writestr(item, docx.read(item.filename))

def process_all_elements(tree, replacements, namespaces):
    # Process all types of elements
    for element in tree.xpath('.//*[self::w:p or self::w:tc or self::w:r or self::w:t or self::w:b]', namespaces=namespaces):
        if element.tag.endswith('tc'):  # Table cell
            for p in element.findall('.//w:p', namespaces):
                process_paragraph(p, replacements, namespaces, is_table_cell=True)
        elif element.tag.endswith('p'):  # Regular paragraph
         process_paragraph(element, replacements, namespaces, is_table_cell=False)
        elif element.tag.endswith('r'):  # Run element
        # Process run elements if needed
            pass
        elif element.tag.endswith('t'):  # Text element
        # Process text elements if needed
            pass
        elif element.tag.endswith('b'):  # Bold element
        # Process bold elements if needed
            pass

def process_paragraph(paragraph, replacements, namespaces, is_table_cell=False):
    runs = paragraph.findall('.//w:r', namespaces)

    # Build a map of the original text and formatting
    text_map = []
    full_text = ''

    for run in runs:
        text_elem = run.find('w:t', namespaces)
        if text_elem is not None and text_elem.text:
            text_map.append({
                'text': text_elem.text,
                'rPr': run.find('w:rPr', namespaces),
                'run': run,
                'parent': run.getparent()
            })
            full_text += text_elem.text

    if not full_text or not text_map:
        return

    # Apply replacements while preserving spaces
    modified_text = full_text
    for key, value in replacements.items():
        # Add spaces around replacements if necessary
        modified_text = modified_text.replace(key, value)

    # Remove all existing runs in the paragraph
    for item in text_map:
        if item['run'].getparent() is not None:
            item['run'].getparent().remove(item['run'])

    # Create new runs with inherited formatting (excluding highlights)
    parent = text_map[0]['parent']
    if parent is None:
        return

    current_pos = 0
    while current_pos < len(modified_text):
        for item in text_map:
            if current_pos >= len(modified_text):
                break

            # Extract a chunk of text matching the length of the original run
            chunk_length = len(item['text'])
            chunk = modified_text[current_pos:current_pos + chunk_length]
            current_pos += chunk_length

            # Skip empty chunks
            if not chunk.strip():
                continue

            # Create a new run
            new_run = etree.SubElement(parent, f'{{{namespaces["w"]}}}r')

            # Copy formatting from the original run, excluding highlights
            if item['rPr'] is not None:
                format_copy = etree.fromstring(etree.tostring(item['rPr']))
                highlight = format_copy.find('.//w:highlight', namespaces)
                if highlight is not None:
                    format_copy.remove(highlight)
                new_run.append(format_copy)

            # Add the text to the new run
            text_elem = etree.SubElement(new_run, f'{{{namespaces["w"]}}}t')
            text_elem.text = chunk

            # Preserve spaces explicitly
            if chunk.startswith(' ') or chunk.endswith(' '):
                text_elem.set(f'{{{namespaces["w"]}}}space', 'preserve')

    # Handle any remaining text
    if current_pos < len(modified_text):
        remaining_text = modified_text[current_pos:]
        if remaining_text.strip():
            new_run = etree.SubElement(parent, f'{{{namespaces["w"]}}}r')
            text_elem = etree.SubElement(new_run, f'{{{namespaces["w"]}}}t')
            text_elem.text = remaining_text
            if remaining_text.startswith(' ') or remaining_text.endswith(' '):
                text_elem.set(f'{{{namespaces["w"]}}}space', 'preserve')


def main():
    # Create database session
    Session = sessionmaker(bind=engine)
    session = Session()

    try:
        # Query database
        stmt = select(datos_table)
        result = session.execute(stmt).first()

        if result:
            # Convert result to dictionary safely
            if hasattr(result, '_mapping'):
                registro_dict = dict(result._mapping)
            elif isinstance(result, dict):
                registro_dict = result
            else:
                registro_dict = dict(zip(datos_table.columns.keys(), result))

            replacements = {
                "#Interviniente_Nombre#": str(registro_dict.get("Interviniente_Nombre", "")),
                "#Agencia#": str(registro_dict.get("Agencia", "")),
                "#Telefono#": str(registro_dict.get("Telefono", "")),
                "#HorarioAgencia#": str(registro_dict.get("HorarioAgencia", "")),
                "#Direccion_Via#": str(registro_dict.get("Direccion_Via", "")),
                "#Direccion_CodigoPostal#": str(registro_dict.get("Direccion_CodigoPostal", "")),
                "#Direccion_Localidad#": str(registro_dict.get("Direccion_Localidad", "")),
                "#Direccion_DenomProvincia#": str(registro_dict.get("Direccion_DenomProvincia", "")),
                "#dia#": str(registro_dict.get("dia", "")),
                "#mes#": str(registro_dict.get("mes", "")),
                "#ano#": str(registro_dict.get("ano", "")),
                "#Propietario#": str(registro_dict.get("Propietario", "")),
                "#Expediente#": str(registro_dict.get("Expediente", "")),
                "#Deuda#": str(registro_dict.get("Deuda", "")),
                "#TextoCesion#": str(registro_dict.get("TextoCesion", "")),
                "#Saldo#": str(registro_dict.get("Saldo", "")),
                "#notariocedente#": str(registro_dict.get("notariocedente", "")),
                "#FechaCesionCedente#": str(registro_dict.get("FechaCesionCedente", "")),
                "#ProtocoloCedente#": str(registro_dict.get("ProtocoloCedente", "")),
                "#Exp_Saldo_TEXTO#": str(registro_dict.get("Exp_Saldo_TEXTO", "")),
                "#SumaSaldoCapitalAgrupadas#": str(registro_dict.get("SumaSaldoCapitalAgrupadas", "")),
                "#CIReferencia#": str(registro_dict.get("CIReferencia", "")),
                "#cccPropietario#": str(registro_dict.get("cccPropietario", "")),
                "#FechaMasVeintiunDias#": str(registro_dict.get("FechaMasVeintiunDias", "")),
            }

            process_document(docx_path, output_path, replacements)
    finally:
        session.close()

if __name__ == "__main__":
    main()