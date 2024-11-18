import pandas as pd
from sqlalchemy import create_engine, Table, Column, String, MetaData

# Configura la conexión SQL (en este caso, SQLite)
engine = create_engine('sqlite:////Users/paulalaorga/Development/cartas/database.db')
metadata = MetaData()

# Define la estructura de la tabla
tabla_datos = Table(
    'datos', metadata,
    Column('ID', String),
    Column('FuenteDocumento', String),
    Column('NumExpediente', String),
    Column('Interviniente_Nombre', String),
    Column('Direccion_Via', String),
    Column('Direccion_CodigoPostal', String),
    Column('Direccion_Localidad', String),
    Column('Direccion_DenomProvincia', String),
    Column('Agencia', String),
    Column('Telefono', String),
    Column('HorarioAgencia', String),
    Column('MailCabeceraCapital', String),
    Column('dia', String),
    Column('mes', String),
    Column('año', String),
    Column('Propietario', String),
    Column('Contrato', String),
    Column('Expediente', String),
    Column('Deuda', String),
    Column('TextoCesion', String),
    Column('Saldo', String),
    Column('notariocedente', String),
    Column('FechaCesionCedente', String),
    Column('ProtocoloCedente', String),
    Column('Exp_Saldo_TEXTO', String),
    Column('SumaSaldoCapitalAgrupadas', String),
    Column('FechaMasVeintiunDias', String),
    Column('cccPropietario', String),
    Column('CIReferencia', String)
)

# Crea la tabla en la base de datos si no existe
metadata.create_all(engine)

# Lee los datos desde output.csv
try:
    df = pd.read_csv('/Users/paulalaorga/Development/cartas/Raw Data/output.csv', encoding='utf-8', delimiter='|')

    # Inserta los datos directamente en la base de datos
    df.to_sql('datos', con=engine, if_exists='replace', index=False)

    print("Datos insertados exitosamente en SQL.")
except Exception as e:
    print(f"Error al insertar datos: {e}")
